use std::cell::Cell;
use std::cell::RefCell;
use std::fs::File;
use std::io::{self, Cursor, Read as _, Seek, SeekFrom, Write};
use std::path::Path;
use std::rc::Rc;

use calamine::{Data, Dimensions, Reader, Xlsx, open_workbook};
use chrono::NaiveDate;
use easyexcel_core::{
    BigDecimal, ClientAnchorData, CoordinateData, ImageData, ImageType, IntoExcelCell,
    WriteCellData,
};
use tempfile::tempdir;
use zip::ZipArchive;

use super::*;

fn test_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

struct FaultyWrite {
    fail_write_at: Option<usize>,
    fail_flush: bool,
    writes: usize,
}

impl FaultyWrite {
    const fn writing(fail_at: usize) -> Self {
        Self {
            fail_write_at: Some(fail_at),
            fail_flush: false,
            writes: 0,
        }
    }

    const fn flushing() -> Self {
        Self {
            fail_write_at: None,
            fail_flush: true,
            writes: 0,
        }
    }
}

impl Write for FaultyWrite {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let call = self.writes;
        self.writes += 1;
        if self.fail_write_at == Some(call) {
            return Err(io::Error::other("injected CSV write failure"));
        }
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.fail_flush {
            return Err(io::Error::other("injected CSV flush failure"));
        }
        Ok(())
    }
}

#[derive(Default)]
struct FailThirdFlush {
    flushes: usize,
}

impl Write for FailThirdFlush {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        let flush = self.flushes;
        self.flushes += 1;
        if flush == 2 {
            Err(io::Error::other("injected CSV finish flush failure"))
        } else {
            Ok(())
        }
    }
}

#[derive(Default)]
struct FailSecondFlush {
    flushes: usize,
}

impl Write for FailSecondFlush {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        let flush = self.flushes;
        self.flushes += 1;
        if flush == 1 {
            Err(io::Error::other("injected CSV into-inner failure"))
        } else {
            Ok(())
        }
    }
}

struct LimitedCursor {
    inner: Cursor<Vec<u8>>,
    max_len: u64,
}

impl LimitedCursor {
    const fn new(max_len: u64) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            max_len,
        }
    }
}

impl std::io::Read for LimitedCursor {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        self.inner.read(buffer)
    }
}

impl Write for LimitedCursor {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let end = self
            .inner
            .position()
            .saturating_add(u64::try_from(buffer.len()).unwrap_or(u64::MAX));
        if end > self.max_len {
            return Err(io::Error::other("injected encrypted output failure"));
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for LimitedCursor {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.inner.seek(position)
    }
}

fn zip_entry(path: &Path, name: &str) -> Result<String> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file).map_err(test_error)?;
    let mut entry = archive.by_name(name).map_err(test_error)?;
    let mut value = String::new();
    entry.read_to_string(&mut value)?;
    Ok(value)
}

fn zip_names(path: &Path) -> Result<Vec<String>> {
    let file = File::open(path)?;
    let mut archive = ZipArchive::new(file).map_err(test_error)?;
    (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .map(|entry| entry.name().to_owned())
                .map_err(test_error)
        })
        .collect::<Result<Vec<_>>>()
}

fn cell_style_id(sheet_xml: &str, cell: &str) -> Option<String> {
    let marker = format!("<c r=\"{cell}\" s=\"");
    sheet_xml
        .split_once(&marker)
        .and_then(|(_, value)| value.split_once('"'))
        .map(|(style, _)| style.to_owned())
}

fn sheet_column_width(sheet_xml: &str, one_based_column: u16) -> Result<f64> {
    let marker = format!("<col min=\"{one_based_column}\"");
    let (_, column) = sheet_xml
        .split_once(&marker)
        .ok_or_else(|| test_error(format!("missing column {one_based_column}")))?;
    let (_, width) = column
        .split_once("width=\"")
        .ok_or_else(|| test_error("missing column width"))?;
    let (width, _) = width
        .split_once('"')
        .ok_or_else(|| test_error("unterminated column width"))?;
    width.parse::<f64>().map_err(test_error)
}

fn sheet_row_height(sheet_xml: &str, one_based_row: u32) -> Result<f64> {
    let marker = format!("<row r=\"{one_based_row}\"");
    let (_, row) = sheet_xml
        .split_once(&marker)
        .ok_or_else(|| test_error(format!("missing row {one_based_row}")))?;
    let (row, _) = row
        .split_once('>')
        .ok_or_else(|| test_error("unterminated row"))?;
    let (_, height) = row
        .split_once("ht=\"")
        .ok_or_else(|| test_error("missing row height"))?;
    let (height, _) = height
        .split_once('"')
        .ok_or_else(|| test_error("unterminated row height"))?;
    height.parse::<f64>().map_err(test_error)
}

#[derive(Clone)]
struct EveryCell {
    cells: Vec<CellValue>,
    fail: bool,
}

thread_local! {
    static USE_WIDE_SCHEMA: Cell<bool> = const { Cell::new(false) };
    static USE_ANNOTATED_WIDE_SCHEMA: Cell<bool> = const { Cell::new(false) };
    static USE_BACKEND_WIDE_SCHEMA: Cell<bool> = const { Cell::new(false) };
}

const TEST_COLUMN: ExcelColumn = ExcelColumn::new("value", "Value", Some(0), 0, None);

struct SparseRow;

struct DimensionRow;

struct StyledAnnotationRow;

impl ExcelRow for StyledAnnotationRow {
    fn schema() -> &'static [ExcelColumn] {
        const FIELD_HEAD_STYLE: ExcelCellStyle = ExcelCellStyle {
            fill_pattern: Some(ExcelFillPattern::Solid),
            fill_foreground_color: Some(ExcelColor::Indexed(14)),
            horizontal_alignment: Some(ExcelHorizontalAlignment::Left),
            ..ExcelCellStyle::new()
        };
        const FIELD_HEAD_FONT: ExcelFontStyle = ExcelFontStyle {
            font_height_in_points: Some(40.0),
            color: Some(ExcelColor::Indexed(51)),
            ..ExcelFontStyle::new()
        };
        const FIELD_CONTENT_STYLE: ExcelCellStyle = ExcelCellStyle {
            fill_pattern: Some(ExcelFillPattern::Solid),
            fill_foreground_color: Some(ExcelColor::Indexed(40)),
            ..ExcelCellStyle::new()
        };
        const FIELD_CONTENT_FONT: ExcelFontStyle = ExcelFontStyle {
            font_height_in_points: Some(50.0),
            color: Some(ExcelColor::Indexed(12)),
            ..ExcelFontStyle::new()
        };
        const COLUMNS: &[ExcelColumn] = &[
            ExcelColumn::new("field", "Field", Some(0), 0, None)
                .with_head_style(FIELD_HEAD_STYLE)
                .with_head_font_style(FIELD_HEAD_FONT)
                .with_content_style(FIELD_CONTENT_STYLE)
                .with_content_font_style(FIELD_CONTENT_FONT),
            ExcelColumn::new("type", "Type", Some(1), 0, None),
        ];
        COLUMNS
    }

    fn write_metadata() -> &'static ExcelWriteMetadata {
        const HEAD_STYLE: ExcelCellStyle = ExcelCellStyle {
            fill_pattern: Some(ExcelFillPattern::Solid),
            fill_foreground_color: Some(ExcelColor::Indexed(10)),
            horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
            ..ExcelCellStyle::new()
        };
        const CONTENT_STYLE: ExcelCellStyle = ExcelCellStyle {
            border_bottom: Some(ExcelBorderStyle::Thin),
            fill_pattern: Some(ExcelFillPattern::Solid),
            fill_foreground_color: Some(ExcelColor::Indexed(17)),
            ..ExcelCellStyle::new()
        };
        const HEAD_FONT: ExcelFontStyle = ExcelFontStyle {
            bold: Some(false),
            font_height_in_points: Some(20.0),
            color: Some(ExcelColor::Indexed(15)),
            ..ExcelFontStyle::new()
        };
        const CONTENT_FONT: ExcelFontStyle = ExcelFontStyle {
            font_height_in_points: Some(30.0),
            color: Some(ExcelColor::Indexed(22)),
            ..ExcelFontStyle::new()
        };
        const METADATA: ExcelWriteMetadata = ExcelWriteMetadata::new()
            .head_style(HEAD_STYLE)
            .content_style(CONTENT_STYLE)
            .head_font_style(HEAD_FONT)
            .content_font_style(CONTENT_FONT);
        &METADATA
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self)
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![
            CellValue::String("field".to_owned()),
            CellValue::String("type".to_owned()),
        ])
    }
}

impl ExcelRow for DimensionRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[
            ExcelColumn::new("field", "Field", Some(0), 0, None).with_column_width(30),
            ExcelColumn::new("type", "Type", Some(1), 0, None),
            ExcelColumn::new("explicit", "Explicit", Some(2), 0, None),
        ];
        COLUMNS
    }

    fn write_metadata() -> &'static ExcelWriteMetadata {
        const METADATA: ExcelWriteMetadata = ExcelWriteMetadata::new()
            .column_width(18)
            .head_row_height(24)
            .content_row_height(16);
        &METADATA
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self)
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![
            CellValue::String("field".to_owned()),
            CellValue::String("type".to_owned()),
            CellValue::String("explicit".to_owned()),
        ])
    }
}

impl ExcelRow for SparseRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("value", "Value", Some(10_000), 0, None)];
        COLUMNS
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self)
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![CellValue::String("value".to_owned())])
    }
}

struct AnchoredImageRow {
    cell: WriteCellData,
}

impl ExcelRow for AnchoredImageRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("cell", "Images", Some(0), 0, None).with_column_width(20)];
        COLUMNS
    }

    fn write_metadata() -> &'static ExcelWriteMetadata {
        const METADATA: ExcelWriteMetadata = ExcelWriteMetadata::new()
            .head_row_height(18)
            .content_row_height(30);
        &METADATA
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self {
            cell: WriteCellData::new(CellValue::Empty),
        })
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![self.cell.to_excel_cell(
            &easyexcel_core::ConvertContext {
                sheet_name: "Images".to_owned(),
                row_index: 1,
                column_index: Some(0),
                field: "cell",
                format: None,
            },
        )?])
    }
}

struct RichTextRow {
    value: RichTextStringData,
}

impl ExcelRow for RichTextRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Rich", Some(0), 0, None)];
        COLUMNS
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self {
            value: RichTextStringData::default(),
        })
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![CellValue::RichText(self.value.clone())])
    }
}

impl ExcelRow for EveryCell {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[
            ExcelColumn::new("empty", "Empty", Some(0), 0, None),
            ExcelColumn::new("string", "String", Some(1), 0, None),
            ExcelColumn::new("error", "Error", Some(2), 0, None),
            ExcelColumn::new("boolean", "Boolean", Some(3), 0, None),
            ExcelColumn::new("integer", "Integer", Some(4), 0, None),
            ExcelColumn::new("float", "Float", Some(5), 0, None),
            ExcelColumn::new("date", "Date", Some(6), 0, Some("%d/%m/%Y")),
            ExcelColumn::new(
                "datetime",
                "DateTime",
                Some(7),
                0,
                Some("%Y-%m-%d %H:%M:%S"),
            ),
            ExcelColumn::new("large", "Large", Some(8), 0, None),
            ExcelColumn::new("missing", "Missing", Some(9), 0, None),
            ExcelColumn::new("formula", "Formula", Some(10), 0, None),
            ExcelColumn::new("link", "Link", Some(11), 0, None),
            ExcelColumn::new("comment", "Comment", Some(12), 0, None),
            ExcelColumn::new("image", "Image", Some(13), 0, None),
            ExcelColumn::new("decimal", "Decimal", Some(14), 0, None),
        ];
        const WIDE_COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("wide", "Wide", Some(65_536), 0, None)];
        const ANNOTATED_WIDE_COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("wide", "Wide", Some(65_536), 0, None).with_column_width(10)];
        const BACKEND_WIDE_COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("wide", "Wide", Some(65_535), 0, None).with_column_width(10)];
        USE_BACKEND_WIDE_SCHEMA.with(|backend_wide| {
            if backend_wide.get() {
                BACKEND_WIDE_COLUMNS
            } else {
                USE_ANNOTATED_WIDE_SCHEMA.with(|annotated_wide| {
                    if annotated_wide.get() {
                        ANNOTATED_WIDE_COLUMNS
                    } else {
                        USE_WIDE_SCHEMA.with(|wide| if wide.get() { WIDE_COLUMNS } else { COLUMNS })
                    }
                })
            }
        })
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Err(ExcelError::Unsupported("test-only writer row".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        if self.fail {
            return Err(ExcelError::Format("row conversion failed".to_owned()));
        }
        Ok(self.cells.clone())
    }
}

fn every_cell() -> EveryCell {
    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    EveryCell {
        cells: vec![
            CellValue::Empty,
            CellValue::String("text".to_owned()),
            CellValue::Error("#DIV/0!".to_owned()),
            CellValue::Bool(true),
            CellValue::Int(-12),
            CellValue::Float(1.25),
            CellValue::Date(date),
            CellValue::DateTime(date.and_hms_opt(12, 34, 56).expect("valid time")),
            CellValue::Int(i64::MAX),
            CellValue::Empty,
            CellValue::Formula("SUM(E2:F2)".to_owned()),
            CellValue::Hyperlink {
                url: "https://www.rust-lang.org".to_owned(),
                text: "Rust".to_owned(),
            },
            CellValue::Comment {
                value: Box::new(CellValue::String("annotated".to_owned())),
                text: "cell note".to_owned(),
            },
            CellValue::Image(tiny_png()),
            CellValue::Decimal("123.45".parse().expect("valid decimal")),
        ],
        fail: false,
    }
}

fn tiny_png() -> Vec<u8> {
    vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
        0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0xf0, 0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00,
        0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]
}

struct RecordingHandler {
    order: i32,
    events: Rc<RefCell<Vec<String>>>,
}

impl WriteHandler for RecordingHandler {
    fn order(&self) -> i32 {
        self.order
    }

    fn before_workbook(&mut self, context: &WriteWorkbookContext) -> Result<()> {
        self.events.borrow_mut().push(format!(
            "{}:before_workbook:{}",
            self.order,
            context.path().display()
        ));
        Ok(())
    }

    fn after_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        self.events
            .borrow_mut()
            .push(format!("{}:after_workbook", self.order));
        Ok(())
    }

    fn before_sheet(&mut self, context: &WriteSheetContext) -> Result<()> {
        self.events.borrow_mut().push(format!(
            "{}:before_sheet:{}",
            self.order,
            context.sheet_name()
        ));
        Ok(())
    }

    fn after_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        self.events
            .borrow_mut()
            .push(format!("{}:after_sheet", self.order));
        Ok(())
    }

    fn before_row(&mut self, context: &WriteRowContext) -> Result<()> {
        self.events
            .borrow_mut()
            .push(format!("{}:before_row:{}", self.order, context.is_head));
        Ok(())
    }

    fn after_row(&mut self, context: &WriteRowContext) -> Result<()> {
        self.events
            .borrow_mut()
            .push(format!("{}:after_row:{}", self.order, context.is_head));
        Ok(())
    }

    fn before_cell(&mut self, context: &mut WriteCellContext) -> Result<()> {
        self.events.borrow_mut().push(format!(
            "{}:before_cell:{}:{}",
            self.order, context.is_head, context.column_index
        ));
        if self.order < 0 {
            match (context.is_head, context.field) {
                (true, Some("empty")) | (false, Some("error")) => context.skip = true,
                (true, Some("string")) => context.value = CellValue::Bool(true),
                (true, Some("error")) => {
                    context.value = CellValue::Error("header-error".to_owned());
                }
                (false, Some("string")) => {
                    context.value = CellValue::String("transformed".to_owned());
                }
                _ => {}
            }
        }
        Ok(())
    }

    fn after_cell(&mut self, context: &WriteCellContext) -> Result<()> {
        self.events.borrow_mut().push(format!(
            "{}:after_cell:{}:{}",
            self.order, context.is_head, context.skip
        ));
        Ok(())
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FailureStage {
    BeforeWorkbook,
    BeforeSheet,
    BeforeHeadRow,
    BeforeHeadCell,
    AfterHeadCell,
    AfterHeadRow,
    BeforeDataRow,
    BeforeDataCell,
    AfterDataCell,
    AfterDataRow,
    AfterSheet,
    AfterWorkbook,
}

struct FailingHandler(FailureStage);

struct InvalidHeaderValueHandler;

impl WriteHandler for InvalidHeaderValueHandler {
    fn before_cell(&mut self, context: &mut WriteCellContext) -> Result<()> {
        context.column_index = u16::MAX;
        context.value = CellValue::Bool(true);
        Ok(())
    }
}

impl FailingHandler {
    fn result(&self, stage: FailureStage) -> Result<()> {
        if self.0 == stage {
            Err(ExcelError::Format("handler failed".to_owned()))
        } else {
            Ok(())
        }
    }
}

impl WriteHandler for FailingHandler {
    fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        self.result(FailureStage::BeforeWorkbook)
    }

    fn after_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        self.result(FailureStage::AfterWorkbook)
    }

    fn before_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        self.result(FailureStage::BeforeSheet)
    }

    fn after_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        self.result(FailureStage::AfterSheet)
    }

    fn before_row(&mut self, context: &WriteRowContext) -> Result<()> {
        self.result(if context.is_head {
            FailureStage::BeforeHeadRow
        } else {
            FailureStage::BeforeDataRow
        })
    }

    fn after_row(&mut self, context: &WriteRowContext) -> Result<()> {
        self.result(if context.is_head {
            FailureStage::AfterHeadRow
        } else {
            FailureStage::AfterDataRow
        })
    }

    fn before_cell(&mut self, context: &mut WriteCellContext) -> Result<()> {
        self.result(if context.is_head {
            FailureStage::BeforeHeadCell
        } else {
            FailureStage::BeforeDataCell
        })
    }

    fn after_cell(&mut self, context: &WriteCellContext) -> Result<()> {
        self.result(if context.is_head {
            FailureStage::AfterHeadCell
        } else {
            FailureStage::AfterDataCell
        })
    }
}

#[test]
fn default_options_and_helpers_are_deterministic() {
    assert_eq!(
        WriteOptions::default(),
        WriteOptions {
            sheet_name: "Sheet1".to_owned(),
            sheet_index: None,
            constant_memory: false,
            need_head: true,
            freeze_head: false,
            freeze_panes: None,
            include_column_indexes: None,
            include_column_field_names: None,
            exclude_column_indexes: Vec::new(),
            exclude_column_field_names: Vec::new(),
            order_by_include_column: false,
            merge_ranges: Vec::new(),
            auto_width: false,
            column_widths: Vec::new(),
            head_style: CellStyle::new().bold(true),
            content_styles: Vec::new(),
            loop_merges: Vec::new(),
            dynamic_head: None,
            password: None,
            charset: CsvCharset::default(),
            with_bom: true,
            converters: ConverterRegistry::default(),
        }
    );
    assert_eq!(excel_date_format(None, "yyyy-mm-dd"), "yyyy-mm-dd");
    assert_eq!(
        excel_date_format(Some("%Y/%m/%d %H:%M:%S"), "unused"),
        "yyyy/mm/dd hh:mm:ss"
    );
    assert_eq!(to_column(0).expect("column"), 0);
    assert_eq!(to_column(usize::from(u16::MAX)).expect("column"), u16::MAX);
    assert!(to_column(usize::from(u16::MAX) + 1).is_err());
    assert_eq!(
        format_error("broken").to_string(),
        "excel format error: broken"
    );
    assert!(LoopMergeStrategy::new(0, 1, 0).is_err());
    assert!(LoopMergeStrategy::new(1, 0, 0).is_err());
    assert!(LoopMergeStrategy::new(1, 1, 0).is_err());
    let strategy = LoopMergeStrategy::new(2, 3, 4).expect("loop merge");
    assert_eq!(strategy.each_rows(), 2);
    assert_eq!(strategy.column_extend(), 3);
    assert_eq!(strategy.column_index(), 4);
    let dynamic = WriteSheet::<EveryCell>::new("Dynamic").head([["User", "Name"], ["User", "Age"]]);
    assert_eq!(
        dynamic.options().dynamic_head,
        Some(vec![
            vec!["User".to_owned(), "Name".to_owned()],
            vec!["User".to_owned(), "Age".to_owned()],
        ])
    );
    let indexed = WriteSheet::<EveryCell>::new_index(5);
    assert_eq!(indexed.options().sheet_index, Some(5));
    assert_eq!(indexed.options().sheet_name, "5");
    let indexed_name = WriteSheet::<EveryCell>::new("Named").sheet_index(6);
    assert_eq!(indexed_name.options().sheet_index, Some(6));
    assert_eq!(indexed_name.options().sheet_name, "Named");
}

#[test]
#[allow(clippy::too_many_lines)]
fn style_model_maps_every_alignment_and_cycles_content_rows() -> Result<()> {
    for (alignment, expected) in [
        (HorizontalAlignment::General, FormatAlign::General),
        (HorizontalAlignment::Left, FormatAlign::Left),
        (HorizontalAlignment::Center, FormatAlign::Center),
        (HorizontalAlignment::Right, FormatAlign::Right),
        (HorizontalAlignment::Fill, FormatAlign::Fill),
        (HorizontalAlignment::Justify, FormatAlign::Justify),
        (HorizontalAlignment::CenterAcross, FormatAlign::CenterAcross),
    ] {
        assert_eq!(horizontal_format_align(alignment), expected);
    }
    for (alignment, expected) in [
        (VerticalAlignment::Top, FormatAlign::Top),
        (VerticalAlignment::Center, FormatAlign::VerticalCenter),
        (VerticalAlignment::Bottom, FormatAlign::Bottom),
        (VerticalAlignment::Justify, FormatAlign::VerticalJustify),
        (
            VerticalAlignment::Distributed,
            FormatAlign::VerticalDistributed,
        ),
    ] {
        assert_eq!(vertical_format_align(alignment), expected);
    }
    for (alignment, expected) in [
        (ExcelHorizontalAlignment::General, FormatAlign::General),
        (ExcelHorizontalAlignment::Left, FormatAlign::Left),
        (ExcelHorizontalAlignment::Center, FormatAlign::Center),
        (ExcelHorizontalAlignment::Right, FormatAlign::Right),
        (ExcelHorizontalAlignment::Fill, FormatAlign::Fill),
        (ExcelHorizontalAlignment::Justify, FormatAlign::Justify),
        (
            ExcelHorizontalAlignment::CenterAcross,
            FormatAlign::CenterAcross,
        ),
        (
            ExcelHorizontalAlignment::Distributed,
            FormatAlign::Distributed,
        ),
    ] {
        assert_eq!(annotation_horizontal_format_align(alignment), expected);
    }
    for (alignment, expected) in [
        (ExcelVerticalAlignment::Top, FormatAlign::Top),
        (ExcelVerticalAlignment::Center, FormatAlign::VerticalCenter),
        (ExcelVerticalAlignment::Bottom, FormatAlign::Bottom),
        (
            ExcelVerticalAlignment::Justify,
            FormatAlign::VerticalJustify,
        ),
        (
            ExcelVerticalAlignment::Distributed,
            FormatAlign::VerticalDistributed,
        ),
    ] {
        assert_eq!(annotation_vertical_format_align(alignment), expected);
    }
    for (border, expected) in [
        (ExcelBorderStyle::None, FormatBorder::None),
        (ExcelBorderStyle::Thin, FormatBorder::Thin),
        (ExcelBorderStyle::Medium, FormatBorder::Medium),
        (ExcelBorderStyle::Dashed, FormatBorder::Dashed),
        (ExcelBorderStyle::Dotted, FormatBorder::Dotted),
        (ExcelBorderStyle::Thick, FormatBorder::Thick),
        (ExcelBorderStyle::Double, FormatBorder::Double),
        (ExcelBorderStyle::Hair, FormatBorder::Hair),
        (ExcelBorderStyle::MediumDashed, FormatBorder::MediumDashed),
        (ExcelBorderStyle::DashDot, FormatBorder::DashDot),
        (ExcelBorderStyle::MediumDashDot, FormatBorder::MediumDashDot),
        (ExcelBorderStyle::DashDotDot, FormatBorder::DashDotDot),
        (
            ExcelBorderStyle::MediumDashDotDot,
            FormatBorder::MediumDashDotDot,
        ),
        (ExcelBorderStyle::SlantDashDot, FormatBorder::SlantDashDot),
    ] {
        assert_eq!(annotation_border_style(border), expected);
    }
    for (pattern, expected) in [
        (ExcelFillPattern::None, FormatPattern::None),
        (ExcelFillPattern::Solid, FormatPattern::Solid),
        (ExcelFillPattern::MediumGray, FormatPattern::MediumGray),
        (ExcelFillPattern::DarkGray, FormatPattern::DarkGray),
        (ExcelFillPattern::LightGray, FormatPattern::LightGray),
        (
            ExcelFillPattern::DarkHorizontal,
            FormatPattern::DarkHorizontal,
        ),
        (ExcelFillPattern::DarkVertical, FormatPattern::DarkVertical),
        (ExcelFillPattern::DarkDown, FormatPattern::DarkDown),
        (ExcelFillPattern::DarkUp, FormatPattern::DarkUp),
        (ExcelFillPattern::DarkGrid, FormatPattern::DarkGrid),
        (ExcelFillPattern::DarkTrellis, FormatPattern::DarkTrellis),
        (
            ExcelFillPattern::LightHorizontal,
            FormatPattern::LightHorizontal,
        ),
        (
            ExcelFillPattern::LightVertical,
            FormatPattern::LightVertical,
        ),
        (ExcelFillPattern::LightDown, FormatPattern::LightDown),
        (ExcelFillPattern::LightUp, FormatPattern::LightUp),
        (ExcelFillPattern::LightGrid, FormatPattern::LightGrid),
        (ExcelFillPattern::LightTrellis, FormatPattern::LightTrellis),
        (ExcelFillPattern::Gray125, FormatPattern::Gray125),
        (ExcelFillPattern::Gray0625, FormatPattern::Gray0625),
    ] {
        assert_eq!(annotation_fill_pattern(pattern), expected);
    }
    for (underline, expected) in [
        (ExcelUnderline::None, FormatUnderline::None),
        (ExcelUnderline::Single, FormatUnderline::Single),
        (ExcelUnderline::Double, FormatUnderline::Double),
        (
            ExcelUnderline::SingleAccounting,
            FormatUnderline::SingleAccounting,
        ),
        (
            ExcelUnderline::DoubleAccounting,
            FormatUnderline::DoubleAccounting,
        ),
    ] {
        assert_eq!(annotation_underline(underline), expected);
    }
    for (script, expected) in [
        (ExcelFontScript::None, FormatScript::None),
        (ExcelFontScript::Superscript, FormatScript::Superscript),
        (ExcelFontScript::Subscript, FormatScript::Subscript),
    ] {
        assert_eq!(annotation_font_script(script), expected);
    }
    assert_eq!(
        annotation_color(ExcelColor::Rgb(0x0012_3456)),
        Color::RGB(0x0012_3456)
    );
    for index in 0..=65 {
        let _ = annotation_color(ExcelColor::Indexed(index));
    }
    assert_eq!(
        annotation_color(ExcelColor::Indexed(10)),
        Color::RGB(0x00ff_0000)
    );
    assert_eq!(annotation_color(ExcelColor::Indexed(64)), Color::Automatic);
    assert_eq!(annotation_color(ExcelColor::Indexed(65)), Color::Default);

    let annotation_cell = ExcelCellStyle {
        hidden: Some(true),
        locked: Some(false),
        quote_prefix: Some(true),
        horizontal_alignment: Some(ExcelHorizontalAlignment::Distributed),
        wrapped: Some(true),
        vertical_alignment: Some(ExcelVerticalAlignment::Distributed),
        rotation: Some(45),
        indent: Some(2),
        border_left: Some(ExcelBorderStyle::Thin),
        border_right: Some(ExcelBorderStyle::Medium),
        border_top: Some(ExcelBorderStyle::Dashed),
        border_bottom: Some(ExcelBorderStyle::Double),
        left_border_color: Some(ExcelColor::Rgb(0x0011_2233)),
        right_border_color: Some(ExcelColor::Rgb(0x0022_3344)),
        top_border_color: Some(ExcelColor::Rgb(0x0033_4455)),
        bottom_border_color: Some(ExcelColor::Rgb(0x0044_5566)),
        fill_pattern: Some(ExcelFillPattern::Solid),
        fill_background_color: Some(ExcelColor::Rgb(0x0055_6677)),
        fill_foreground_color: Some(ExcelColor::Rgb(0x0066_7788)),
        shrink_to_fit: Some(true),
        data_format: Some(ExcelDataFormat::Custom("0.00")),
    };
    assert_ne!(
        apply_annotation_cell_style(Format::new(), annotation_cell),
        Format::new()
    );
    assert_ne!(
        apply_annotation_cell_style(
            Format::new(),
            ExcelCellStyle {
                data_format: Some(ExcelDataFormat::Builtin(14)),
                ..ExcelCellStyle::new()
            }
        ),
        Format::new()
    );
    let disabled_cell = ExcelCellStyle {
        hidden: Some(false),
        locked: Some(true),
        quote_prefix: Some(false),
        wrapped: Some(false),
        shrink_to_fit: Some(false),
        ..ExcelCellStyle::new()
    };
    let _ = apply_annotation_cell_style(Format::new(), disabled_cell);

    let annotation_font = ExcelFontStyle {
        font_name: Some("Arial"),
        font_height_in_points: Some(12.5),
        italic: Some(true),
        strikeout: Some(true),
        color: Some(ExcelColor::Rgb(0x0077_8899)),
        type_offset: Some(ExcelFontScript::Superscript),
        underline: Some(ExcelUnderline::DoubleAccounting),
        charset: Some(1),
        bold: Some(true),
    };
    assert_ne!(
        apply_annotation_font_style(Format::new(), annotation_font),
        Format::new()
    );
    let disabled_font = ExcelFontStyle {
        italic: Some(false),
        strikeout: Some(false),
        bold: Some(false),
        ..ExcelFontStyle::new()
    };
    let _ = apply_annotation_font_style(Format::new(), disabled_font);

    let head_style = CellStyle::new()
        .bold(true)
        .italic(true)
        .font_color(0x00ff_0000)
        .background_color(0x0000_ff00)
        .horizontal_alignment(HorizontalAlignment::Center)
        .vertical_alignment(VerticalAlignment::Center)
        .wrap_text(true)
        .number_format("0.00");
    let content_styles = vec![
        CellStyle::new().font_color(0x0000_00ff),
        CellStyle::new().font_color(0x00ff_0000),
    ];
    let directory = tempdir()?;
    let path = directory.path().join("styles.xlsx");
    write_xlsx::<EveryCell, _>(
        &path,
        &WriteOptions {
            head_style,
            content_styles,
            ..WriteOptions::default()
        },
        vec![every_cell(), every_cell()],
    )?;

    let styles = zip_entry(&path, "xl/styles.xml")?;
    assert!(styles.contains("<b/>"));
    assert!(styles.contains("<i/>"));
    assert!(styles.contains("rgb=\"FFFF0000\""));
    assert!(styles.contains("rgb=\"FF00FF00\""));
    assert!(styles.contains("formatCode=\"0.00\""));
    assert!(styles.contains("horizontal=\"center\""));
    assert!(styles.contains("vertical=\"center\""));
    assert!(styles.contains("wrapText=\"1\""));
    let sheet = zip_entry(&path, "xl/worksheets/sheet1.xml")?;
    assert_ne!(
        cell_style_id(&sheet, "A2").expect("first content style"),
        cell_style_id(&sheet, "A3").expect("second content style")
    );
    Ok(())
}

#[test]
fn annotation_dimensions_apply_field_type_and_explicit_precedence() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("annotation-dimensions.xlsx");
    write_xlsx::<DimensionRow, _>(
        &path,
        &WriteOptions {
            column_widths: vec![(2, 40)],
            ..WriteOptions::default()
        },
        vec![DimensionRow],
    )?;

    let sheet = zip_entry(&path, "xl/worksheets/sheet1.xml")?;
    let field_width = sheet_column_width(&sheet, 1)?;
    let type_width = sheet_column_width(&sheet, 2)?;
    let explicit_width = sheet_column_width(&sheet, 3)?;
    assert!((field_width - type_width - 12.0).abs() < f64::EPSILON);
    assert!((explicit_width - type_width - 22.0).abs() < f64::EPSILON);
    assert!((sheet_row_height(&sheet, 1)? - 24.0).abs() < f64::EPSILON);
    assert!((sheet_row_height(&sheet, 2)? - 16.0).abs() <= 0.25);
    Ok(())
}

#[test]
fn annotation_styles_apply_field_type_and_builder_precedence() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("annotation-styles.xlsx");
    write_xlsx::<StyledAnnotationRow, _>(
        &path,
        &WriteOptions {
            head_style: CellStyle::new().bold(true).font_color(0x0000_ff00),
            ..WriteOptions::default()
        },
        vec![StyledAnnotationRow],
    )?;

    let styles = zip_entry(&path, "xl/styles.xml")?;
    assert!(styles.contains("rgb=\"FF00FF00\""));
    assert!(styles.contains("rgb=\"FFFF0000\""));
    assert!(styles.contains("rgb=\"FF00CCFF\""));
    assert!(styles.contains("rgb=\"FF008000\""));
    assert!(styles.contains("rgb=\"FF0000FF\""));
    assert!(styles.contains("rgb=\"FFC0C0C0\""));
    assert!(styles.contains("<sz val=\"50\"/>"));
    assert!(styles.contains("style=\"thin\""));

    let sheet = zip_entry(&path, "xl/worksheets/sheet1.xml")?;
    assert_ne!(cell_style_id(&sheet, "A1"), cell_style_id(&sheet, "B1"));
    assert_ne!(cell_style_id(&sheet, "A2"), cell_style_id(&sheet, "B2"));

    let java_path = directory.path().join("java-indexed-annotation-styles.xlsx");
    write_xlsx::<StyledAnnotationRow, _>(
        &java_path,
        &WriteOptions::default(),
        vec![StyledAnnotationRow],
    )?;
    let java_styles = zip_entry(&java_path, "xl/styles.xml")?;
    for expected in [
        "rgb=\"FFFF00FF\"",
        "rgb=\"FFFFCC00\"",
        "rgb=\"FF00CCFF\"",
        "rgb=\"FF0000FF\"",
        "rgb=\"FFFF0000\"",
        "rgb=\"FF00FFFF\"",
        "rgb=\"FF008000\"",
        "rgb=\"FFC0C0C0\"",
    ] {
        assert!(java_styles.contains(expected), "missing {expected}");
    }
    for expected in [20, 30, 40, 50] {
        assert!(java_styles.contains(&format!("<sz val=\"{expected}\"/>")));
    }
    Ok(())
}

#[test]
fn dynamic_multi_level_head_merges_parents_and_offsets_data_rows() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("dynamic-head.xlsx");
    let options = WriteOptions {
        sheet_name: "Dynamic".to_owned(),
        include_column_indexes: Some(vec![0, 1, 2]),
        dynamic_head: Some(vec![
            vec!["User".to_owned(), "Empty".to_owned()],
            vec!["User".to_owned(), "String".to_owned()],
            vec!["Meta".to_owned()],
        ]),
        freeze_head: true,
        ..WriteOptions::default()
    };
    assert_eq!(dynamic_head_rows(&options)?, 2);
    write_xlsx::<EveryCell, _>(&path, &options, vec![every_cell()])?;

    let mut workbook: Xlsx<_> = open_workbook(&path).map_err(test_error)?;
    let range = workbook.worksheet_range("Dynamic").map_err(test_error)?;
    assert_eq!(
        range.get_value((0, 0)),
        Some(&Data::String("User".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 1)),
        Some(&Data::String("String".to_owned()))
    );
    assert_eq!(
        range.get_value((2, 1)),
        Some(&Data::String("text".to_owned()))
    );
    assert_eq!(
        workbook
            .merge_cells_by_sheet_name("Dynamic")
            .map_err(test_error)?,
        vec![Dimensions::new((0, 0), (0, 1))]
    );
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn dynamic_head_validation_and_backend_failures_are_typed() -> Result<()> {
    let directory = tempdir()?;
    assert_eq!(head_level_to_row(0)?, 0);
    assert!(head_level_to_row(usize::MAX).is_err());
    assert_eq!(
        dynamic_head_rows(&WriteOptions {
            need_head: false,
            dynamic_head: Some(Vec::new()),
            ..WriteOptions::default()
        })?,
        0
    );
    assert!(
        dynamic_head_rows(&WriteOptions {
            dynamic_head: Some(Vec::new()),
            ..WriteOptions::default()
        })
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("empty-head.xlsx"),
            &WriteOptions {
                dynamic_head: Some(Vec::new()),
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("empty-head-paths.xlsx"),
            &WriteOptions {
                dynamic_head: Some(vec![Vec::new(); EveryCell::schema().len()]),
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    let invalid_head_options = WriteOptions {
        dynamic_head: Some(vec![Vec::new(); EveryCell::schema().len()]),
        ..WriteOptions::default()
    };
    let mut workbook = Workbook::new();
    assert!(
        append_rows_to_worksheet::<EveryCell, _>(
            workbook.add_worksheet(),
            &invalid_head_options,
            Vec::new(),
            &mut [],
            WriteProgress {
                next_row: 0,
                next_data_index: 0,
            },
            true,
            EveryCell::write_metadata(),
        )
        .is_err()
    );
    let invalid_head_height = ExcelWriteMetadata::new().head_row_height(16);
    assert!(
        append_rows_to_worksheet::<EveryCell, _>(
            workbook.add_worksheet(),
            &WriteOptions {
                include_column_indexes: Some(Vec::new()),
                ..WriteOptions::default()
            },
            Vec::new(),
            &mut [],
            WriteProgress {
                next_row: 1_048_576,
                next_data_index: 0,
            },
            true,
            &invalid_head_height,
        )
        .is_err()
    );
    let invalid_content_height = ExcelWriteMetadata::new().content_row_height(16);
    assert!(
        append_rows_to_worksheet::<EveryCell, _>(
            workbook.add_worksheet(),
            &WriteOptions {
                need_head: false,
                ..WriteOptions::default()
            },
            vec![every_cell()],
            &mut [],
            WriteProgress {
                next_row: 1_048_576,
                next_data_index: 0,
            },
            true,
            &invalid_content_height,
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("mismatched-head.xlsx"),
            &WriteOptions {
                include_column_indexes: Some(vec![0, 1]),
                dynamic_head: Some(vec![vec!["Only one".to_owned()]]),
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    USE_WIDE_SCHEMA.with(|wide| wide.set(true));
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("wide-dynamic-head.xlsx"),
            &WriteOptions {
                dynamic_head: Some(vec![vec!["Wide".to_owned()]]),
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    USE_WIDE_SCHEMA.with(|wide| wide.set(false));
    USE_ANNOTATED_WIDE_SCHEMA.with(|wide| wide.set(true));
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("annotated-wide-column.xlsx"),
            &WriteOptions::default(),
            Vec::new(),
        )
        .is_err()
    );
    USE_ANNOTATED_WIDE_SCHEMA.with(|wide| wide.set(false));
    USE_BACKEND_WIDE_SCHEMA.with(|wide| wide.set(true));
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("backend-wide-column.xlsx"),
            &WriteOptions::default(),
            Vec::new(),
        )
        .is_err()
    );
    USE_BACKEND_WIDE_SCHEMA.with(|wide| wide.set(false));

    let head = vec![vec!["Group".to_owned()], vec!["Group".to_owned()]];
    for columns in [
        vec![(65_536, 0, &TEST_COLUMN), (65_537, 0, &TEST_COLUMN)],
        vec![(65_535, 0, &TEST_COLUMN), (65_536, 0, &TEST_COLUMN)],
        vec![(16_383, 0, &TEST_COLUMN), (16_384, 0, &TEST_COLUMN)],
    ] {
        let mut raw = Workbook::new();
        let worksheet = raw.add_worksheet();
        assert!(
            merge_dynamic_head_groups(
                worksheet,
                &columns,
                &head,
                SheetStyleContext::head(&CellStyle::default(), &ExcelWriteMetadata::new()),
            )
            .is_err()
        );
    }
    assert!(
        dynamic_head_rows(&WriteOptions {
            dynamic_head: Some(vec![Vec::new()]),
            ..WriteOptions::default()
        })
        .is_err()
    );
    assert!(!same_dynamic_head_group(
        &[
            vec!["A".to_owned(), "X".to_owned()],
            vec!["B".to_owned(), "X".to_owned()]
        ],
        0,
        1,
        1
    ));
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn columns_are_ordered_by_physical_index_order_and_schema_position() {
    const SCHEMA: &[ExcelColumn] = &[
        ExcelColumn::new("third", "Third", Some(2), 0, None),
        ExcelColumn::new("second", "Second", Some(1), 5, None),
        ExcelColumn::new("first", "First", Some(1), 1, None),
        ExcelColumn::new("implicit", "Implicit", None, 0, None),
    ];
    let actual = ordered_columns(SCHEMA)
        .into_iter()
        .map(|(physical, schema, column)| (physical, schema, column.field))
        .collect::<Vec<_>>();
    assert_eq!(
        actual,
        vec![
            (1, 2, "first"),
            (1, 1, "second"),
            (2, 0, "third"),
            (3, 3, "implicit")
        ]
    );

    let by_index = selected_columns(
        SCHEMA,
        &WriteOptions {
            include_column_indexes: Some(vec![2, 1]),
            order_by_include_column: true,
            ..WriteOptions::default()
        },
    );
    assert_eq!(
        by_index
            .iter()
            .map(|(_, _, column)| column.field)
            .collect::<Vec<_>>(),
        vec!["third", "first", "second"]
    );
    assert_eq!(
        by_index
            .iter()
            .map(|(physical, _, _)| *physical)
            .collect::<Vec<_>>(),
        vec![0, 1, 2]
    );

    let by_name = selected_columns(
        SCHEMA,
        &WriteOptions {
            include_column_field_names: Some(vec!["implicit".to_owned(), "first".to_owned()]),
            order_by_include_column: true,
            ..WriteOptions::default()
        },
    );
    assert_eq!(
        by_name
            .iter()
            .map(|(_, _, column)| column.field)
            .collect::<Vec<_>>(),
        vec!["implicit", "first"]
    );

    let excluded = selected_columns(
        SCHEMA,
        &WriteOptions {
            exclude_column_indexes: vec![2],
            exclude_column_field_names: vec!["second".to_owned()],
            ..WriteOptions::default()
        },
    );
    assert_eq!(
        excluded
            .iter()
            .map(|(_, _, column)| column.field)
            .collect::<Vec<_>>(),
        vec!["first", "implicit"]
    );

    let dynamic = selected_columns(
        easyexcel_core::DynamicRow::schema(),
        &WriteOptions {
            dynamic_head: Some(vec![
                vec!["First".to_owned()],
                vec!["Second".to_owned()],
                vec!["Third".to_owned()],
            ]),
            include_column_indexes: Some(vec![2, 0]),
            exclude_column_indexes: vec![1],
            order_by_include_column: true,
            ..WriteOptions::default()
        },
    );
    assert_eq!(
        dynamic
            .iter()
            .map(|(physical, source, column)| (*physical, *source, column.field))
            .collect::<Vec<_>>(),
        vec![(0, 2, ""), (1, 0, "")]
    );
    assert!(
        selected_dynamic_columns(
            3,
            &WriteOptions {
                include_column_field_names: Some(vec!["unknown".to_owned()]),
                ..WriteOptions::default()
            }
        )
        .is_empty()
    );
    assert_eq!(
        selected_dynamic_columns(
            2,
            &WriteOptions {
                order_by_include_column: true,
                ..WriteOptions::default()
            }
        )
        .iter()
        .map(|(physical, source, _)| (*physical, *source))
        .collect::<Vec<_>>(),
        vec![(0, 0), (1, 1)]
    );
}

#[test]
fn dynamic_row_layout_omits_a_synthetic_head_and_accepts_a_dynamic_head() -> Result<()> {
    let options = WriteOptions::default();
    assert_eq!(head_rows_for_schema_state(true, &options)?, 0);
    assert!(dynamic_columns_for_row(true, 3, &options).is_some());
    assert!(dynamic_columns_for_row(false, 3, &options).is_none());

    let headed_options = WriteOptions {
        dynamic_head: Some(vec![
            vec!["Name".to_owned()],
            vec!["Unused".to_owned()],
            vec!["Score".to_owned()],
        ]),
        ..WriteOptions::default()
    };
    assert_eq!(head_rows_for_schema_state(true, &headed_options)?, 1);
    assert!(dynamic_columns_for_row(true, 3, &headed_options).is_none());

    let mut writer = create_csv_record_writer(Box::new(Vec::<u8>::new()), &options.charset, true)?;
    let mut rows = [Ok(vec![
        CellValue::String("Alice".to_owned()),
        CellValue::Empty,
        CellValue::Int(7),
    ])]
    .into_iter();
    let progress = append_csv_records(
        &mut writer,
        &options,
        &[],
        true,
        &mut rows,
        &mut [],
        0,
        0,
        true,
    )?;
    assert_eq!(progress.next_row, 1);
    assert_eq!(progress.next_data_index, 1);
    finish_csv_record_writer(writer)
}

#[test]
fn writer_emits_headers_and_every_supported_cell_type() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("all.xlsx");
    write_xlsx::<EveryCell, _>(
        &path,
        &WriteOptions {
            sheet_name: "Values".to_owned(),
            constant_memory: false,
            need_head: true,
            freeze_head: true,
            freeze_panes: None,
            ..WriteOptions::default()
        },
        vec![every_cell()],
    )?;

    let mut workbook: Xlsx<_> = open_workbook(&path).map_err(test_error)?;
    let range = workbook.worksheet_range("Values").map_err(test_error)?;
    assert_eq!(
        range.get_value((0, 1)),
        Some(&Data::String("String".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 1)),
        Some(&Data::String("text".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 2)),
        Some(&Data::String("#DIV/0!".to_owned()))
    );
    assert_eq!(range.get_value((1, 3)), Some(&Data::Bool(true)));
    assert_eq!(range.get_value((1, 4)), Some(&Data::Float(-12.0)));
    assert_eq!(range.get_value((1, 5)), Some(&Data::Float(1.25)));
    assert!(matches!(range.get_value((1, 6)), Some(Data::DateTime(_))));
    assert!(matches!(range.get_value((1, 7)), Some(Data::DateTime(_))));
    assert_eq!(
        range.get_value((1, 8)),
        Some(&Data::String(i64::MAX.to_string()))
    );
    assert_eq!(range.get_value((1, 9)), Some(&Data::Empty));
    assert_eq!(
        range.get_value((1, 11)),
        Some(&Data::String("Rust".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 12)),
        Some(&Data::String("annotated".to_owned()))
    );
    assert_eq!(range.get_value((1, 14)), Some(&Data::Float(123.45)));
    let formulas = workbook.worksheet_formula("Values").map_err(test_error)?;
    assert!(
        formulas
            .get_value((1, 10))
            .is_some_and(|formula| formula.contains("SUM(E2:F2)"))
    );

    let sheet = zip_entry(&path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains("<hyperlink ref=\"L2\""));
    let comments = zip_entry(&path, "xl/comments1.xml")?;
    assert!(comments.contains("cell note"));
    let names = zip_names(&path)?;
    assert!(names.iter().any(|name| name == "xl/media/image1.png"));
    Ok(())
}

#[test]
fn write_cell_data_emits_multiple_images_with_java_anchor_semantics() -> Result<()> {
    let bytes = tiny_png();
    let spanning = ClientAnchorData::new()
        .coordinates(
            CoordinateData::new()
                .relative_last_row_index(1)
                .relative_last_column_index(1),
        )
        .left(5)
        .top(6)
        .right(7)
        .bottom(8)
        .anchor_type(AnchorType::MoveDontResize);
    let zero_absolute_defers = ClientAnchorData::new()
        .coordinates(
            CoordinateData::new()
                .first_row_index(0)
                .first_column_index(0)
                .relative_first_row_index(1)
                .relative_first_column_index(1)
                .relative_last_row_index(1)
                .relative_last_column_index(1),
        )
        .anchor_type(AnchorType::DontMoveDoResize);
    let absolute = ClientAnchorData::new()
        .coordinates(
            CoordinateData::new()
                .first_row_index(4)
                .first_column_index(3)
                .last_row_index(4)
                .last_column_index(3),
        )
        .anchor_type(AnchorType::DontMoveAndResize);
    let cell = WriteCellData::new(CellValue::String("caption".to_owned())).image_data_list([
        ImageData::new(bytes.clone()).image_type(ImageType::Png),
        ImageData::new(bytes.clone()).anchor(spanning),
        ImageData::new(bytes.clone()).anchor(zero_absolute_defers),
        ImageData::new(bytes).anchor(absolute),
    ]);
    let directory = tempdir()?;
    let path = directory.path().join("multiple-images.xlsx");
    write_xlsx::<AnchoredImageRow, _>(
        &path,
        &WriteOptions {
            sheet_name: "Images".to_owned(),
            column_widths: vec![(1, 12)],
            ..WriteOptions::default()
        },
        [AnchoredImageRow { cell }],
    )?;

    let mut workbook: Xlsx<_> = open_workbook(&path).map_err(test_error)?;
    let range = workbook.worksheet_range("Images").map_err(test_error)?;
    assert_eq!(
        range.get_value((1, 0)),
        Some(&Data::String("caption".to_owned()))
    );
    let drawing = zip_entry(&path, "xl/drawings/drawing1.xml")?;
    assert_eq!(drawing.matches("<xdr:twoCellAnchor").count(), 4);
    assert_eq!(drawing.matches("editAs=\"oneCell\"").count(), 2);
    assert_eq!(drawing.matches("editAs=\"absolute\"").count(), 1);
    assert!(drawing.contains("<xdr:col>3</xdr:col>"));
    assert!(drawing.contains("<xdr:row>4</xdr:row>"));
    Ok(())
}

#[test]
fn rich_text_writer_applies_java_whole_and_utf16_interval_fonts() -> Result<()> {
    let whole = WriteFont::new()
        .font_name("Aptos")
        .font_height_in_points(13.0)
        .italic(true)
        .strikeout(true)
        .color(ExcelColor::Indexed(10))
        .type_offset(ExcelFontScript::Subscript)
        .underline(ExcelUnderline::Single)
        .charset(1)
        .bold(true);
    let override_font = WriteFont::new()
        .italic(false)
        .strikeout(false)
        .color(ExcelColor::Rgb(0x00_80_00))
        .type_offset(ExcelFontScript::None)
        .underline(ExcelUnderline::None)
        .bold(false);
    let rich = RichTextStringData::new("A😀BC")
        .apply_font(whole)
        .apply_font_range(1, 3, override_font.clone())
        .apply_font_range(3, 5, WriteFont::new().color(ExcelColor::Indexed(11)))
        .apply_font_range(4, 5, override_font);
    let runs = rich_text_runs(&rich)?;
    assert_eq!(
        runs.iter()
            .map(|(_, text)| text.as_str())
            .collect::<Vec<_>>(),
        ["A", "😀", "B", "C"]
    );
    assert_eq!(utf16_byte_index("A😀BC", 0), Some(0));
    assert_eq!(utf16_byte_index("A😀BC", 1), Some(1));
    assert_eq!(utf16_byte_index("A😀BC", 2), None);
    assert_eq!(utf16_byte_index("A😀BC", 5), Some("A😀BC".len()));
    assert_eq!(utf16_byte_index("A😀BC", 6), None);
    assert_eq!(rich_text_runs(&RichTextStringData::new("plain"))?.len(), 1);
    assert!(
        rich_text_runs(&RichTextStringData::new("abc").apply_font_range(1, 1, WriteFont::new()))
            .is_err()
    );
    assert!(
        rich_text_runs(&RichTextStringData::new("abc").apply_font_range(0, 4, WriteFont::new()))
            .is_err()
    );
    assert!(
        rich_text_runs(&RichTextStringData::new("😀").apply_font_range(0, 1, WriteFont::new()))
            .is_err()
    );
    let _ = rich_text_format(&WriteFont::new());

    let directory = tempdir()?;
    let path = directory.path().join("rich-text.xlsx");
    write_xlsx::<RichTextRow, _>(
        &path,
        &WriteOptions::default(),
        [
            RichTextRow {
                value: rich.clone(),
            },
            RichTextRow {
                value: RichTextStringData::new(""),
            },
        ],
    )?;
    let shared_strings = zip_entry(&path, "xl/sharedStrings.xml")?;
    assert!(shared_strings.contains("<t>A</t>"));
    assert!(shared_strings.contains("<t>😀</t>"));
    assert!(shared_strings.contains("rgb=\"FF008000\""));
    assert!(shared_strings.contains("<vertAlign val=\"subscript\"/>"));

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let invalid = RichTextStringData::new("abc").apply_font_range(2, 2, WriteFont::new());
    assert!(write_rich_text(worksheet, 0, 0, &invalid, &Format::new()).is_err());
    let metadata = ExcelWriteMetadata::new();
    assert!(
        write_cell(
            worksheet,
            0,
            0,
            &TEST_COLUMN,
            &CellValue::RichText(invalid),
            SheetStyleContext::content(None, &metadata).column(&TEST_COLUMN),
            &ImageLayout::default(),
        )
        .is_err()
    );
    assert!(
        write_rich_text(
            worksheet,
            u32::MAX,
            0,
            &RichTextStringData::new(""),
            &Format::new(),
        )
        .is_err()
    );
    assert!(write_rich_text(worksheet, u32::MAX, 0, &rich, &Format::new()).is_err());
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn image_anchor_layout_and_validation_cover_java_coordinate_boundaries() -> Result<()> {
    let columns = selected_columns(AnchoredImageRow::schema(), &WriteOptions::default());
    let layout = ImageLayout::new(
        &columns,
        &WriteOptions {
            column_widths: vec![(0, 12), (2, 0)],
            ..WriteOptions::default()
        },
        AnchoredImageRow::write_metadata(),
        1,
    )?;
    assert_eq!(layout.column_width(0), 89);
    assert_eq!(layout.column_width(1), 64);
    assert_eq!(layout.column_width(2), 0);
    assert_eq!(layout.row_height(0), 24);
    assert_eq!(layout.row_height(1), 40);
    assert_eq!(excel_column_width_pixels(0), 0);
    assert_eq!(excel_row_height_pixels(None), 20);
    assert_eq!(resolve_anchor_coordinate(4, Some(3), Some(8), "row")?, 3);
    assert_eq!(resolve_anchor_coordinate(4, Some(0), Some(-2), "row")?, 2);
    assert_eq!(resolve_anchor_coordinate(4, None, None, "row")?, 4);
    assert!(resolve_anchor_coordinate(0, None, Some(-1), "row").is_err());

    let bytes = tiny_png();
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let invalid_anchors = [
        ClientAnchorData::new().coordinates(CoordinateData::new().relative_first_row_index(-1)),
        ClientAnchorData::new().coordinates(CoordinateData::new().relative_first_column_index(-1)),
        ClientAnchorData::new().coordinates(CoordinateData::new().relative_last_row_index(-1)),
        ClientAnchorData::new().coordinates(CoordinateData::new().relative_last_column_index(-1)),
        ClientAnchorData::new()
            .coordinates(CoordinateData::new().first_row_index(2).last_row_index(1)),
        ClientAnchorData::new().coordinates(
            CoordinateData::new()
                .relative_first_column_index(70_000)
                .relative_last_column_index(70_000),
        ),
        ClientAnchorData::new()
            .coordinates(CoordinateData::new().relative_last_column_index(70_000)),
        ClientAnchorData::new().coordinates(
            CoordinateData::new()
                .first_row_index(1_048_576)
                .last_row_index(1_048_576),
        ),
        ClientAnchorData::new().left(64),
        ClientAnchorData::new().top(20),
    ];
    for anchor in invalid_anchors {
        assert!(
            insert_image_data(
                worksheet,
                0,
                0,
                &ImageData::new(bytes.clone()).anchor(anchor),
                &ImageLayout::default(),
            )
            .is_err()
        );
    }
    assert!(
        insert_image_data(
            worksheet,
            0,
            0,
            &ImageData::new([1, 2, 3]),
            &ImageLayout::default(),
        )
        .is_err()
    );

    let width_overflow = ImageLayout {
        column_widths: HashMap::from([(0, u32::MAX)]),
        ..ImageLayout::default()
    };
    let two_columns =
        ClientAnchorData::new().coordinates(CoordinateData::new().relative_last_column_index(1));
    assert!(
        insert_image_data(
            worksheet,
            0,
            0,
            &ImageData::new(bytes.clone()).anchor(two_columns),
            &width_overflow,
        )
        .is_err()
    );
    let height_overflow = ImageLayout {
        content_row_height: u32::MAX,
        ..ImageLayout::default()
    };
    let two_rows =
        ClientAnchorData::new().coordinates(CoordinateData::new().relative_last_row_index(1));
    assert!(
        insert_image_data(
            worksheet,
            0,
            0,
            &ImageData::new(bytes.clone()).anchor(two_rows),
            &height_overflow,
        )
        .is_err()
    );
    let valid_image = image_from_buffer(&bytes)?;
    assert!(insert_scaled_image(worksheet, u32::MAX, 0, &valid_image, 0, 0).is_err());
    let metadata = ExcelWriteMetadata::new();
    let style = SheetStyleContext::content(None, &metadata).column(&TEST_COLUMN);
    assert!(
        write_cell(
            worksheet,
            0,
            0,
            &TEST_COLUMN,
            &CellValue::Images {
                value: Box::new(CellValue::Decimal(
                    "9".repeat(400).parse().expect("valid huge decimal"),
                )),
                images: Vec::new(),
            },
            style,
            &ImageLayout::default(),
        )
        .is_err()
    );
    assert!(
        write_cell(
            worksheet,
            0,
            0,
            &TEST_COLUMN,
            &CellValue::Images {
                value: Box::new(CellValue::Empty),
                images: vec![ImageData::new([1, 2, 3])],
            },
            style,
            &ImageLayout::default(),
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn decimal_writer_rejects_values_outside_xlsx_numeric_range() {
    let huge: BigDecimal = "9".repeat(400).parse().expect("valid large decimal");
    let metadata = ExcelWriteMetadata::new();
    let style = SheetStyleContext::content(None, &metadata).column(&TEST_COLUMN);
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    assert!(
        write_cell(
            worksheet,
            0,
            0,
            &TEST_COLUMN,
            &CellValue::Decimal(huge),
            style,
            &ImageLayout::default(),
        )
        .is_err()
    );
    assert!(
        write_cell(
            worksheet,
            u32::MAX,
            0,
            &TEST_COLUMN,
            &CellValue::Decimal("1.5".parse().expect("valid decimal")),
            style,
            &ImageLayout::default(),
        )
        .is_err()
    );
}

#[test]
fn constant_memory_writer_can_omit_headers_and_freeze_request() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("stream.xlsx");
    write_xlsx::<EveryCell, _>(
        &path,
        &WriteOptions {
            sheet_name: "Stream".to_owned(),
            constant_memory: true,
            need_head: false,
            freeze_head: true,
            freeze_panes: None,
            ..WriteOptions::default()
        },
        vec![every_cell(), every_cell()],
    )?;
    let mut workbook: Xlsx<_> = open_workbook(path).map_err(test_error)?;
    let range = workbook.worksheet_range("Stream").map_err(test_error)?;
    assert_eq!(
        range.get_value((0, 1)),
        Some(&Data::String("text".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 1)),
        Some(&Data::String("text".to_owned()))
    );
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn stateful_writer_supports_multiple_sheets_and_idempotent_finish() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("multi.xlsx");
    let events = Rc::new(RefCell::new(Vec::new()));
    let handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(RecordingHandler {
        order: 5,
        events: Rc::clone(&events),
    })];
    let first = WriteSheet::<EveryCell>::new("Users")
        .sheet_index(7)
        .freeze_head(true)
        .merge_cells(MergeRange::new(0, 0, 0, 1))
        .auto_width(true)
        .column_width(0, 20)
        .head_style(CellStyle::new().italic(true))
        .content_style(CellStyle::new().bold(true))
        .content_styles([CellStyle::new().wrap_text(true)])
        .loop_merge(LoopMergeStrategy::new(2, 1, 0)?);
    let second = WriteSheet::<EveryCell>::new("Archive")
        .sheet_index(9)
        .need_head(false)
        .constant_memory(true);
    assert_eq!(first.options().sheet_name, "Users");
    assert_eq!(first.options().sheet_index, Some(7));
    assert!(first.options().freeze_head);
    assert!(first.options().auto_width);
    assert_eq!(first.options().column_widths, vec![(0, 20)]);
    assert!(first.options().head_style.italic);
    assert_eq!(first.options().content_styles.len(), 1);
    assert!(first.options().content_styles[0].wrap_text);
    assert_eq!(first.options().loop_merges.len(), 1);
    assert!(!second.options().need_head);
    assert!(second.options().constant_memory);

    let mut writer = ExcelWriter::with_handlers(&path, handlers);
    assert!(!writer.is_finished());
    writer
        .write(vec![every_cell(), every_cell()], &first)?
        .write(vec![every_cell(), every_cell()], &first)?
        .write(vec![every_cell(), every_cell()], &second)?;
    writer.write(Vec::new(), &WriteSheet::<EveryCell>::new_index(7))?;
    writer.write(Vec::new(), &WriteSheet::<EveryCell>::new_index(9))?;
    writer.finish()?;
    assert!(writer.is_finished());
    writer.finish()?;
    let Err(error) = writer.write(vec![every_cell()], &first) else {
        panic!("finished writer must reject data");
    };
    assert!(error.to_string().contains("already finished"));

    let actual = events.borrow();
    assert_eq!(
        actual
            .iter()
            .filter(|event| event.contains("before_workbook"))
            .count(),
        1
    );
    assert_eq!(
        actual
            .iter()
            .filter(|event| event.contains("before_sheet"))
            .count(),
        2
    );
    assert_eq!(
        actual
            .iter()
            .filter(|event| event.contains("after_sheet"))
            .count(),
        2
    );
    assert_eq!(
        actual
            .iter()
            .filter(|event| event.contains("after_workbook"))
            .count(),
        1
    );
    drop(actual);

    let mut workbook: Xlsx<_> = open_workbook(path).map_err(test_error)?;
    assert_eq!(workbook.sheet_names(), vec!["Users", "Archive"]);
    assert_eq!(
        workbook
            .merge_cells_by_sheet_name("Users")
            .map_err(test_error)?,
        vec![
            Dimensions::new((0, 0), (0, 1)),
            Dimensions::new((1, 0), (2, 0)),
            Dimensions::new((3, 0), (4, 0)),
        ]
    );
    let users = workbook.worksheet_range("Users").map_err(test_error)?;
    assert_eq!(
        users.get_value((1, 1)),
        Some(&Data::String("text".to_owned()))
    );
    assert_eq!(
        users.get_value((4, 1)),
        Some(&Data::String("text".to_owned()))
    );
    let archive = workbook.worksheet_range("Archive").map_err(test_error)?;
    assert_eq!(
        archive.get_value((0, 1)),
        Some(&Data::String("text".to_owned()))
    );
    assert_eq!(
        archive.get_value((1, 1)),
        Some(&Data::String("text".to_owned()))
    );
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn stateful_writer_propagates_start_sheet_and_finish_failures() -> Result<()> {
    let directory = tempdir()?;
    let sheet = WriteSheet::<EveryCell>::new("Values");

    let handlers: Vec<Box<dyn WriteHandler>> =
        vec![Box::new(FailingHandler(FailureStage::BeforeWorkbook))];
    let mut rejected = ExcelWriter::with_handlers(directory.path().join("rejected.xlsx"), handlers);
    assert!(rejected.write(Vec::new(), &sheet).is_err());

    let handlers: Vec<Box<dyn WriteHandler>> =
        vec![Box::new(FailingHandler(FailureStage::BeforeWorkbook))];
    let mut rejected_finish =
        ExcelWriter::with_handlers(directory.path().join("rejected-finish.xlsx"), handlers);
    assert!(rejected_finish.finish().is_err());

    let handlers: Vec<Box<dyn WriteHandler>> =
        vec![Box::new(FailingHandler(FailureStage::AfterWorkbook))];
    let mut rejected_after =
        ExcelWriter::with_handlers(directory.path().join("rejected-after.xlsx"), handlers);
    rejected_after.write(Vec::new(), &sheet)?;
    assert!(rejected_after.finish().is_err());

    let mut schema_change = ExcelWriter::new(directory.path().join("schema-change.xlsx"));
    schema_change.write(Vec::new(), &sheet)?;
    USE_WIDE_SCHEMA.with(|wide| wide.set(true));
    let schema_change_result = schema_change.write(Vec::new(), &sheet);
    USE_WIDE_SCHEMA.with(|wide| wide.set(false));
    assert!(matches!(schema_change_result, Err(ExcelError::Format(_))));

    let invalid = WriteSheet::<EveryCell>::new("bad/name");
    let mut invalid_sheet = ExcelWriter::new(directory.path().join("invalid.xlsx"));
    assert!(invalid_sheet.write(Vec::new(), &invalid).is_err());

    let mut invalid_output = ExcelWriter::new(directory.path());
    assert!(invalid_output.finish().is_err());

    let mut csv = ExcelWriter::with_handlers_and_options(
        directory.path().join("invalid-charset.CSV"),
        Vec::new(),
        WriteOptions {
            charset: CsvCharset::new("not-a-charset"),
            ..WriteOptions::default()
        },
    );
    assert!(matches!(
        csv.write(Vec::new(), &sheet),
        Err(ExcelError::Unsupported(_))
    ));
    let mut protected_csv = ExcelWriter::with_handlers_and_password(
        directory.path().join("protected.csv"),
        Vec::new(),
        Some("secret".to_owned()),
    );
    assert!(matches!(
        protected_csv.finish(),
        Err(ExcelError::Unsupported(_))
    ));
    assert!(matches!(
        validate_csv_options(&WriteOptions {
            password: Some("secret".to_owned()),
            ..WriteOptions::default()
        }),
        Err(ExcelError::Unsupported(_))
    ));
    let mut xls = ExcelWriter::new(directory.path().join("stateful.XLS"));
    assert!(matches!(xls.finish(), Err(ExcelError::Unsupported(_))));

    let mut failed_xlsx_append = ExcelWriter::new(directory.path().join("failed-xlsx-append.xlsx"));
    failed_xlsx_append.write(vec![every_cell()], &sheet)?;
    let mut broken = every_cell();
    broken.fail = true;
    assert!(matches!(
        failed_xlsx_append.write(vec![broken.clone()], &sheet),
        Err(ExcelError::Format(_))
    ));

    let mut missing_cached_sheet =
        ExcelWriter::new(directory.path().join("missing-cached-sheet.xlsx"));
    missing_cached_sheet.write(Vec::new(), &sheet)?;
    missing_cached_sheet.workbook = Workbook::new();
    assert!(missing_cached_sheet.write(Vec::new(), &sheet).is_err());

    let mut no_autofit = ExcelWriter::new(directory.path().join("no-autofit.xlsx"));
    no_autofit
        .write(Vec::new(), &sheet)?
        .write(Vec::new(), &sheet)?;

    let mut failed_csv_append = ExcelWriter::new(directory.path().join("failed-csv-append.csv"));
    failed_csv_append.write(vec![every_cell()], &sheet)?;
    assert!(matches!(
        failed_csv_append.write(vec![broken], &sheet),
        Err(ExcelError::Format(_))
    ));

    let missing_parent = directory.path().join("missing").join("stateful.csv");
    let mut missing_csv_output = ExcelWriter::new(missing_parent);
    assert!(missing_csv_output.finish().is_err());

    for stage in [FailureStage::BeforeSheet, FailureStage::AfterSheet] {
        let handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(FailingHandler(stage))];
        let mut failed_sheet = ExcelWriter::with_handlers(
            directory
                .path()
                .join(format!("stateful-csv-handler-{}.csv", stage as u8)),
            handlers,
        );
        assert!(failed_sheet.write(Vec::new(), &sheet).is_err());
    }

    let mut failed_csv_finish = ExcelWriter::new(directory.path().join("failed-csv-finish.csv"));
    failed_csv_finish.start()?;
    failed_csv_finish.csv_writer = Some(csv::WriterBuilder::new().from_writer(
        CsvEncodingWriter::new(
            Box::new(FaultyWrite::flushing()),
            CsvEncoding::Standard(encoding_rs::UTF_8),
        ),
    ));
    assert!(failed_csv_finish.finish().is_err());

    let mut missing_csv_writer = None;
    finish_stateful_csv_writer(&mut missing_csv_writer)?;
    Ok(())
}

#[test]
fn stateful_csv_appends_batches_with_one_head_and_one_sheet_lifecycle() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("stateful.csv");
    let events = Rc::new(RefCell::new(Vec::new()));
    let handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(RecordingHandler {
        order: 1,
        events: Rc::clone(&events),
    })];
    let options = WriteOptions {
        charset: CsvCharset::new("GBK"),
        with_bom: false,
        ..WriteOptions::default()
    };
    let sheet = WriteSheet::<EveryCell>::new("Values").sheet_index(3);
    let indexed_alias = WriteSheet::<EveryCell>::new_index(3);
    let mut writer = ExcelWriter::with_handlers_and_options(&path, handlers, options);
    writer
        .write(vec![every_cell()], &sheet)?
        .write(vec![every_cell()], &indexed_alias)?;
    USE_WIDE_SCHEMA.with(|wide| wide.set(true));
    let schema_change_result = writer.write(Vec::new(), &sheet);
    USE_WIDE_SCHEMA.with(|wide| wide.set(false));
    assert!(matches!(schema_change_result, Err(ExcelError::Format(_))));
    let other = WriteSheet::<EveryCell>::new("Other");
    assert!(matches!(
        writer.write(Vec::new(), &other),
        Err(ExcelError::Unsupported(_))
    ));
    writer.finish()?;
    writer.finish()?;

    let bytes = std::fs::read(path)?;
    assert!(!bytes.starts_with(b"\xEF\xBB\xBF"));
    let (decoded, actual, had_errors) = encoding_rs::GBK.decode(&bytes);
    assert_eq!(actual, encoding_rs::GBK);
    assert!(!had_errors);
    let mut csv = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_reader(decoded.as_bytes());
    let records = csv
        .records()
        .collect::<csv::Result<Vec<_>>>()
        .map_err(test_error)?;
    assert_eq!(records.len(), 3);
    assert_eq!(records[0].get(1), Some("String"));
    assert_eq!(records[1].get(1), Some("text"));
    assert_eq!(records[2].get(1), Some("text"));

    let events = events.borrow();
    for event in [
        "before_workbook",
        "after_workbook",
        "before_sheet",
        "after_sheet",
    ] {
        assert_eq!(
            events.iter().filter(|value| value.contains(event)).count(),
            1,
            "event {event}"
        );
    }
    Ok(())
}

#[test]
fn ordered_handlers_observe_transform_and_skip_the_full_lifecycle() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("handled.xlsx");
    let events = Rc::new(RefCell::new(Vec::new()));
    let mut handlers: Vec<Box<dyn WriteHandler>> = vec![
        Box::new(RecordingHandler {
            order: 10,
            events: Rc::clone(&events),
        }),
        Box::new(RecordingHandler {
            order: -10,
            events: Rc::clone(&events),
        }),
    ];
    write_xlsx_with_handlers::<EveryCell, _>(
        &path,
        &WriteOptions::default(),
        vec![every_cell()],
        &mut handlers,
    )?;

    let actual = events.borrow();
    assert!(actual[0].starts_with("-10:before_workbook:"));
    assert!(actual[1].starts_with("10:before_workbook:"));
    assert!(actual.iter().any(|event| event == "-10:after_workbook"));
    assert!(actual.iter().any(|event| event == "10:after_workbook"));
    drop(actual);

    let mut workbook: Xlsx<_> = open_workbook(path).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(range.get_value((0, 0)), None);
    assert_eq!(range.get_value((0, 1)), Some(&Data::Bool(true)));
    assert_eq!(
        range.get_value((1, 1)),
        Some(&Data::String("transformed".to_owned()))
    );
    assert_eq!(range.get_value((1, 2)), Some(&Data::Empty));
    Ok(())
}

#[test]
fn every_handler_failure_stage_is_propagated() -> Result<()> {
    let directory = tempdir()?;
    for (index, stage) in [
        FailureStage::BeforeWorkbook,
        FailureStage::BeforeSheet,
        FailureStage::BeforeHeadRow,
        FailureStage::BeforeHeadCell,
        FailureStage::AfterHeadCell,
        FailureStage::AfterHeadRow,
        FailureStage::BeforeDataRow,
        FailureStage::BeforeDataCell,
        FailureStage::AfterDataCell,
        FailureStage::AfterDataRow,
        FailureStage::AfterSheet,
        FailureStage::AfterWorkbook,
    ]
    .into_iter()
    .enumerate()
    {
        let mut handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(FailingHandler(stage))];
        let error = write_xlsx_with_handlers::<EveryCell, _>(
            &directory.path().join(format!("handler-{index}.xlsx")),
            &WriteOptions::default(),
            vec![every_cell()],
            &mut handlers,
        )
        .expect_err("handler failure must propagate");
        assert_eq!(error.to_string(), "excel format error: handler failed");
    }

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let mut handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(InvalidHeaderValueHandler)];
    assert!(
        write_headers_with_handlers(
            worksheet,
            &selected_columns(EveryCell::schema(), &WriteOptions::default()),
            "Sheet1",
            SheetStyleContext::head(&CellStyle::default(), &ExcelWriteMetadata::new()),
            &mut handlers,
            &ImageLayout::default(),
        )
        .is_err()
    );
    let worksheet = workbook.add_worksheet();
    let columns = selected_columns(EveryCell::schema(), &WriteOptions::default());
    let head = columns
        .iter()
        .map(|_| vec!["Head".to_owned()])
        .collect::<Vec<_>>();
    let mut handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(InvalidHeaderValueHandler)];
    assert!(
        write_dynamic_headers_with_handlers(
            worksheet,
            &columns,
            &head,
            "Sheet2",
            SheetStyleContext::head(&CellStyle::default(), &ExcelWriteMetadata::new()),
            &mut handlers,
            &ImageLayout::default(),
        )
        .is_err()
    );
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn conversion_configuration_column_and_save_failures_propagate() -> Result<()> {
    let directory = tempdir()?;
    let mut broken = every_cell();
    broken.fail = true;
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("broken.xlsx"),
            &WriteOptions::default(),
            vec![broken]
        )
        .is_err()
    );

    let wide_column = Box::leak(Box::new(ExcelColumn::new(
        "wide",
        "Wide",
        Some(65_536),
        0,
        None,
    )));
    let columns = vec![(65_536, 0, &*wide_column)];
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    assert!(write_headers(worksheet, &columns).is_err());
    assert!(
        write_data_row(
            worksheet,
            0,
            &columns,
            &[CellValue::String("wide".to_owned())]
        )
        .is_err()
    );

    USE_WIDE_SCHEMA.with(|wide| wide.set(true));
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("wide-head.xlsx"),
            &WriteOptions::default(),
            Vec::new()
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("wide-data.xlsx"),
            &WriteOptions {
                need_head: false,
                ..WriteOptions::default()
            },
            vec![every_cell()]
        )
        .is_err()
    );
    USE_WIDE_SCHEMA.with(|wide| wide.set(false));

    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("bad-freeze.xlsx"),
            &WriteOptions {
                freeze_panes: Some((1_048_576, 0)),
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );

    let long_name = Box::leak("x".repeat(32_768).into_boxed_str());
    let long_header = Box::leak(Box::new(ExcelColumn::new(
        "long",
        long_name,
        Some(0),
        0,
        None,
    )));
    assert!(write_headers(worksheet, &[(0, 0, &*long_header)]).is_err());

    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    let invalid_row = 1_048_576;
    for value in [
        CellValue::String("text".to_owned()),
        CellValue::Bool(true),
        CellValue::Int(1),
        CellValue::Int(i64::MAX),
        CellValue::Float(1.0),
        CellValue::Date(date),
        CellValue::DateTime(date.and_hms_opt(1, 2, 3).expect("valid time")),
        CellValue::Formula("1+1".to_owned()),
        CellValue::Hyperlink {
            url: "https://www.rust-lang.org".to_owned(),
            text: "Rust".to_owned(),
        },
        CellValue::Comment {
            value: Box::new(CellValue::String("value".to_owned())),
            text: "note".to_owned(),
        },
        CellValue::Comment {
            value: Box::new(CellValue::Empty),
            text: "note".to_owned(),
        },
        CellValue::Image(tiny_png()),
    ] {
        let metadata = Box::leak(Box::new(ExcelColumn::new(
            "value",
            "Value",
            Some(0),
            0,
            None,
        )));
        assert!(write_data_row(worksheet, invalid_row, &[(0, 0, &*metadata)], &[value]).is_err());
    }
    let metadata = Box::leak(Box::new(ExcelColumn::new(
        "image",
        "Image",
        Some(0),
        0,
        None,
    )));
    for bytes in [vec![1, 2, 3], vec![0; 8]] {
        assert!(
            write_data_row(
                worksheet,
                0,
                &[(0, 0, &*metadata)],
                &[CellValue::Image(bytes)]
            )
            .is_err()
        );
    }
    assert!(
        write_data_row(
            worksheet,
            0,
            &[(0, 0, &*metadata)],
            &[CellValue::Comment {
                value: Box::new(CellValue::String("value".to_owned())),
                text: "x".repeat(32_768),
            }]
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("bad-sheet.xlsx"),
            &WriteOptions {
                sheet_name: "bad/name".to_owned(),
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("bad-merge.xlsx"),
            &WriteOptions {
                merge_ranges: vec![MergeRange::new(0, 0, 0, 0)],
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("bad-width.xlsx"),
            &WriteOptions {
                column_widths: vec![(u16::MAX, 20)],
                ..WriteOptions::default()
            },
            Vec::new()
        )
        .is_err()
    );
    assert!(
        apply_loop_merges(worksheet, u32::MAX, 0, &[LoopMergeStrategy::new(2, 1, 0)?]).is_err()
    );
    assert!(
        apply_loop_merges(worksheet, 0, 0, &[LoopMergeStrategy::new(1, 2, u16::MAX)?]).is_err()
    );
    assert!(apply_loop_merges(worksheet, 0, 0, &[LoopMergeStrategy::new(1, 2, 16_383)?]).is_err());
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("bad-loop-merge.xlsx"),
            &WriteOptions {
                loop_merges: vec![LoopMergeStrategy::new(1, 2, 16_383)?],
                ..WriteOptions::default()
            },
            vec![every_cell()]
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(directory.path(), &WriteOptions::default(), Vec::new()).is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            &directory.path().join("missing").join("encrypted.xlsx"),
            &WriteOptions {
                password: Some("123456".to_owned()),
                ..WriteOptions::default()
            },
            Vec::new(),
        )
        .is_err()
    );
    let mut invalid_encrypted = Workbook::new();
    invalid_encrypted
        .add_worksheet()
        .set_name("Duplicate")
        .map_err(test_error)?;
    invalid_encrypted
        .add_worksheet()
        .set_name("Duplicate")
        .map_err(test_error)?;
    assert!(
        save_workbook(
            &mut invalid_encrypted,
            &directory.path().join("invalid-encrypted.xlsx"),
            Some("123456"),
        )
        .is_err()
    );
    let mut create_failure = Workbook::new();
    create_failure.add_worksheet();
    let mut create_output = LimitedCursor::new(0);
    assert!(
        save_encrypted_workbook_to(&mut create_failure, "123456", &mut create_output,).is_err()
    );
    let mut finalize_failure = Workbook::new();
    finalize_failure.add_worksheet();
    let mut finalize_output = LimitedCursor::new(4_096);
    assert!(
        save_encrypted_workbook_to(&mut finalize_failure, "123456", &mut finalize_output,).is_err()
    );
    let mut successful = Workbook::new();
    successful.add_worksheet();
    let mut successful_output = LimitedCursor::new(u64::MAX);
    save_encrypted_workbook_to(&mut successful, "123456", &mut successful_output)?;
    Ok(())
}

#[test]
fn csv_writer_emits_bom_all_cell_values_and_handler_lifecycle() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("values.csv");
    let events = Rc::new(RefCell::new(Vec::new()));
    let mut handlers: Vec<Box<dyn WriteHandler>> = vec![
        Box::new(RecordingHandler {
            order: 10,
            events: Rc::clone(&events),
        }),
        Box::new(RecordingHandler {
            order: -1,
            events: Rc::clone(&events),
        }),
    ];
    write_csv_with_handlers::<EveryCell, _>(
        &path,
        &WriteOptions::default(),
        [every_cell()],
        &mut handlers,
    )?;
    let bytes = std::fs::read(&path)?;
    assert!(bytes.starts_with(b"\xEF\xBB\xBF"));
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_reader(bytes.as_slice());
    let records = reader
        .records()
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(test_error)?;
    assert_eq!(records.len(), 2);
    assert_eq!(records[0].get(0), Some(""));
    assert_eq!(records[0].get(1), Some("true"));
    assert_eq!(records[0].get(2), Some("header-error"));
    assert_eq!(records[1].get(1), Some("transformed"));
    assert_eq!(records[1].get(2), Some(""));
    assert_eq!(records[1].get(3), Some("true"));
    assert_eq!(records[1].get(13), Some(""));
    assert!(
        events
            .borrow()
            .iter()
            .any(|event| event == "10:after_workbook")
    );
    Ok(())
}

#[test]
fn csv_writer_encodes_java_charsets_and_configurable_bom() -> Result<()> {
    let directory = tempdir()?;
    let mut row = every_cell();
    row.cells[1] = CellValue::String("姓名".to_owned());

    for (name, encoding, expected_bom) in [
        ("utf-8", encoding_rs::UTF_8, b"\xEF\xBB\xBF".as_slice()),
        ("GBK", encoding_rs::GBK, b"".as_slice()),
        ("UTF-16BE", encoding_rs::UTF_16BE, b"\xFE\xFF".as_slice()),
        ("UTF-16LE", encoding_rs::UTF_16LE, b"\xFF\xFE".as_slice()),
    ] {
        let path = directory
            .path()
            .join(format!("{}.csv", name.to_lowercase()));
        write_csv_with_handlers::<EveryCell, _>(
            &path,
            &WriteOptions {
                charset: CsvCharset::new(name),
                ..WriteOptions::default()
            },
            [row.clone()],
            &mut [],
        )?;
        let bytes = std::fs::read(path)?;
        assert!(bytes.starts_with(expected_bom));
        let (decoded, actual_encoding, had_errors) = encoding.decode(&bytes);
        assert_eq!(actual_encoding, encoding);
        assert!(!had_errors);
        assert!(decoded.contains("姓名"));
    }

    let no_bom = directory.path().join("no-bom.csv");
    write_csv_with_handlers::<EveryCell, _>(
        &no_bom,
        &WriteOptions {
            with_bom: false,
            ..WriteOptions::default()
        },
        [row],
        &mut [],
    )?;
    assert!(!std::fs::read(no_bom)?.starts_with(b"\xEF\xBB\xBF"));

    let unsupported = directory.path().join("unsupported.csv");
    let error = write_csv_with_handlers::<EveryCell, _>(
        &unsupported,
        &WriteOptions {
            charset: CsvCharset::new("not-a-charset"),
            ..WriteOptions::default()
        },
        Vec::new(),
        &mut [],
    )
    .expect_err("unknown charset must be rejected");
    assert!(matches!(error, ExcelError::Unsupported(_)));
    assert!(!unsupported.exists());
    Ok(())
}

#[test]
fn csv_transcoding_writer_handles_chunk_boundaries_and_invalid_utf8() -> Result<()> {
    let mut split = CsvEncodingWriter::new(
        Box::new(Vec::<u8>::new()),
        CsvEncoding::Standard(encoding_rs::GBK),
    );
    assert_eq!(split.write(&[0xE5])?, 1);
    assert!(split.finish().is_err());

    let mut split_ok = CsvEncodingWriter::new(
        Box::new(Vec::<u8>::new()),
        CsvEncoding::Standard(encoding_rs::GBK),
    );
    assert_eq!(split_ok.write(&[0xE5])?, 1);
    assert_eq!(split_ok.write(&[0xA7, 0x93])?, 2);
    split_ok.finish()?;

    let mut invalid = CsvEncodingWriter::new(
        Box::new(Vec::<u8>::new()),
        CsvEncoding::Standard(encoding_rs::UTF_8),
    );
    assert!(invalid.write(&[0xFF]).is_err());

    let mut long = CsvEncodingWriter::new(Box::new(Vec::<u8>::new()), CsvEncoding::Utf16Be);
    let value = "姓名".repeat(5_000);
    long.write_all(value.as_bytes())?;
    long.finish()?;

    let mut standard_long = CsvEncodingWriter::new(
        Box::new(Vec::<u8>::new()),
        CsvEncoding::Standard(encoding_rs::GBK),
    );
    standard_long.write_all(value.as_bytes())?;
    standard_long.finish()?;

    let mut failing_utf16 =
        CsvEncodingWriter::new(Box::new(FaultyWrite::writing(0)), CsvEncoding::Utf16Le);
    assert!(failing_utf16.write_all(value.as_bytes()).is_err());

    let mut direct_utf16 = Vec::new();
    CsvEncodingWriter::encode_utf16(&mut direct_utf16, "姓名", u16::to_le_bytes)?;
    assert_eq!(direct_utf16, [0xD3, 0x59, 0x0D, 0x54]);

    let mut finish_failure = CsvEncodingWriter::new(
        Box::new(FaultyWrite::writing(1)),
        CsvEncoding::Standard(encoding_rs::ISO_2022_JP),
    );
    finish_failure.write_all("日本".as_bytes())?;
    assert!(finish_failure.finish().is_err());
    Ok(())
}

#[test]
fn csv_writer_supports_dynamic_heads_no_head_and_configuration_failures() -> Result<()> {
    let directory = tempdir()?;
    let mut dynamic = (0..EveryCell::schema().len())
        .map(|index| vec!["Group".to_owned(), format!("Column {index}")])
        .collect::<Vec<_>>();
    dynamic[0].pop();
    write_csv_with_handlers::<EveryCell, _>(
        &directory.path().join("dynamic.csv"),
        &WriteOptions {
            dynamic_head: Some(dynamic),
            ..WriteOptions::default()
        },
        Vec::new(),
        &mut [],
    )?;
    write_csv_with_handlers::<EveryCell, _>(
        &directory.path().join("no-head.csv"),
        &WriteOptions {
            need_head: false,
            ..WriteOptions::default()
        },
        [every_cell()],
        &mut [],
    )?;
    assert!(
        write_csv_with_handlers::<EveryCell, _>(
            &directory.path().join("bad-head.csv"),
            &WriteOptions {
                dynamic_head: Some(vec![vec!["Only one".to_owned()]]),
                ..WriteOptions::default()
            },
            Vec::new(),
            &mut []
        )
        .is_err()
    );
    assert!(
        write_csv_with_handlers::<EveryCell, _>(
            &directory.path().join("empty-head.csv"),
            &WriteOptions {
                dynamic_head: Some(vec![Vec::new(); EveryCell::schema().len()]),
                ..WriteOptions::default()
            },
            Vec::new(),
            &mut []
        )
        .is_err()
    );
    assert!(
        write_csv_with_handlers::<EveryCell, _>(
            &directory.path().join("conversion.csv"),
            &WriteOptions::default(),
            [EveryCell {
                cells: Vec::new(),
                fail: true
            }],
            &mut []
        )
        .is_err()
    );
    assert!(
        write_csv_with_handlers::<EveryCell, _>(
            directory.path(),
            &WriteOptions::default(),
            Vec::new(),
            &mut []
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn csv_writer_propagates_every_handler_failure() -> Result<()> {
    let directory = tempdir()?;
    for stage in [
        FailureStage::BeforeWorkbook,
        FailureStage::BeforeSheet,
        FailureStage::BeforeHeadRow,
        FailureStage::BeforeHeadCell,
        FailureStage::AfterHeadCell,
        FailureStage::AfterHeadRow,
        FailureStage::BeforeDataRow,
        FailureStage::BeforeDataCell,
        FailureStage::AfterDataCell,
        FailureStage::AfterDataRow,
        FailureStage::AfterSheet,
        FailureStage::AfterWorkbook,
    ] {
        let mut handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(FailingHandler(stage))];
        assert!(
            write_csv_with_handlers::<EveryCell, _>(
                &directory
                    .path()
                    .join(format!("failure-{}.csv", stage as u8)),
                &WriteOptions::default(),
                [every_cell()],
                &mut handlers
            )
            .is_err()
        );
    }
    let mut dynamic_handlers: Vec<Box<dyn WriteHandler>> =
        vec![Box::new(FailingHandler(FailureStage::BeforeHeadRow))];
    assert!(
        write_csv_with_handlers::<EveryCell, _>(
            &directory.path().join("dynamic-handler-failure.csv"),
            &WriteOptions {
                dynamic_head: Some(vec![vec!["Head".to_owned()]; EveryCell::schema().len()]),
                ..WriteOptions::default()
            },
            Vec::new(),
            &mut dynamic_handlers
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn csv_writer_to_owned_stream_validates_options() {
    assert!(
        write_csv_to_writer::<EveryCell, _, _>(
            Path::new("stream.csv"),
            Cursor::new(Vec::new()),
            &WriteOptions::default(),
            [every_cell()],
            &mut [],
        )
        .is_ok()
    );
    assert!(
        write_csv_to_writer::<EveryCell, _, _>(
            Path::new("stream.csv"),
            Cursor::new(Vec::new()),
            &WriteOptions {
                charset: CsvCharset::new("not-a-charset"),
                ..WriteOptions::default()
            },
            [every_cell()],
            &mut [],
        )
        .is_err()
    );
}

#[test]
#[allow(clippy::too_many_lines)]
fn csv_writer_propagates_io_faults_and_column_overflow() {
    let write_errors = (0..64)
        .filter(|fail_at| {
            write_csv_to::<EveryCell, _>(
                Path::new("fault.csv"),
                Box::new(FaultyWrite::writing(*fail_at)),
                &WriteOptions::default(),
                [every_cell()],
                &mut [],
            )
            .is_err()
        })
        .count();
    assert!(write_errors > 0);
    assert!(
        write_csv_to::<EveryCell, _>(
            Path::new("fault.csv"),
            Box::new(FaultyWrite::flushing()),
            &WriteOptions::default(),
            [every_cell()],
            &mut []
        )
        .is_err()
    );
    assert!(
        write_csv_to::<EveryCell, _>(
            Path::new("finish-fault.csv"),
            Box::new(FailThirdFlush::default()),
            &WriteOptions::default(),
            Vec::new(),
            &mut []
        )
        .is_err()
    );
    assert!(
        write_csv_to::<EveryCell, _>(
            Path::new("into-inner-fault.csv"),
            Box::new(FailSecondFlush::default()),
            &WriteOptions::default(),
            Vec::new(),
            &mut []
        )
        .is_err()
    );
    assert!(
        write_csv_to::<EveryCell, _>(
            Path::new("charset-fault.csv"),
            Box::new(Vec::<u8>::new()),
            &WriteOptions {
                charset: CsvCharset::new("not-a-charset"),
                ..WriteOptions::default()
            },
            Vec::new(),
            &mut []
        )
        .is_err()
    );
    for (options, rows) in [
        (WriteOptions::default(), Vec::<SparseRow>::new()),
        (
            WriteOptions {
                dynamic_head: Some(vec![vec!["Dynamic".to_owned()]]),
                ..WriteOptions::default()
            },
            Vec::<SparseRow>::new(),
        ),
        (
            WriteOptions {
                need_head: false,
                ..WriteOptions::default()
            },
            vec![SparseRow],
        ),
    ] {
        assert!(
            write_csv_to::<SparseRow, _>(
                Path::new("record-fault.csv"),
                Box::new(FaultyWrite::writing(1)),
                &options,
                rows,
                &mut [],
            )
            .is_err()
        );
    }

    USE_WIDE_SCHEMA.with(|wide| wide.set(true));
    let wide_result = write_csv_to::<EveryCell, _>(
        Path::new("wide.csv"),
        Box::new(Vec::<u8>::new()),
        &WriteOptions::default(),
        [every_cell()],
        &mut [],
    );
    USE_WIDE_SCHEMA.with(|wide| wide.set(false));
    assert!(wide_result.is_err());
    USE_WIDE_SCHEMA.with(|wide| wide.set(true));
    let wide_data_result = write_csv_to::<EveryCell, _>(
        Path::new("wide-data.csv"),
        Box::new(Vec::<u8>::new()),
        &WriteOptions {
            need_head: false,
            ..WriteOptions::default()
        },
        [every_cell()],
        &mut [],
    );
    USE_WIDE_SCHEMA.with(|wide| wide.set(false));
    assert!(wide_data_result.is_err());
    assert!(csv_record(&[]).is_empty());
}
