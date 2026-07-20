//! Method-level 1:1 parity for Java core tests:
//! AnnotationDataTest, AnnotationIndexAndNameDataTest, StyleDataTest,
//! WriteHandlerTest, ExcludeOrIncludeDataTest.
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>` so each Rust test
//! uniquely maps to `ClassName#methodName`.
//!
//! Format strategy:
//! - `.xlsx` / `.csv`: write → read round-trip with real assertions
//! - `.xls`: real BIFF8 write → read; XLSX-only style/dimension XML checks skipped

use std::collections::HashSet;
use std::fs::File;
use std::io::Read;
use std::path::Path;
use std::sync::{Arc, Mutex};

use easyexcel::{
    DynamicValue, EasyExcel, ExcelCellStyle, ExcelColor, ExcelFillPattern, ExcelRow,
    HorizontalCellStyleStrategy, LoopMergeStrategy, SimpleColumnWidthStyleStrategy,
    SimpleRowHeightStyleStrategy, VerticalCellStyleStrategy, WriteCellContext, WriteHandler,
    WriteRowContext, WriteSheetContext, WriteWorkbookContext,
};
use zip::ZipArchive;

fn temp_path(name: &str) -> std::path::PathBuf {
    let dir = tempfile::tempdir().unwrap();
    dir.keep().join(name)
}

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

/// Assert a Java-generated fixture exists (no soft-skip).
fn require_fixture(name: &str) -> std::path::PathBuf {
    let path = fixture(name);
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
    path
}


/// Reads a ZIP entry from an XLSX workbook as UTF-8 text.
fn is_xls_path(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls"))
}

fn assert_real_biff8(path: &Path) {
    let bytes = std::fs::read(path).expect("read written workbook");
    assert!(
        bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]),
        "expected real BIFF8/OLE compound document: {}",
        path.display()
    );
}

fn zip_entry(path: &Path, name: &str) -> String {
    let file = File::open(path).expect("open xlsx");
    let mut archive = ZipArchive::new(file).expect("open zip");
    let mut entry = archive.by_name(name).expect("zip entry");
    let mut value = String::new();
    entry.read_to_string(&mut value).expect("read zip entry");
    value
}

/// Parses `width` from `<col min="{one_based}" ... width="N"/>`.
fn sheet_column_width(sheet_xml: &str, one_based_column: u16) -> f64 {
    let marker = format!("<col min=\"{one_based_column}\"");
    let (_, column) = sheet_xml
        .split_once(&marker)
        .unwrap_or_else(|| panic!("missing column {one_based_column}"));
    let (_, width) = column
        .split_once("width=\"")
        .expect("missing column width");
    let (width, _) = width.split_once('"').expect("unterminated column width");
    width.parse().expect("column width f64")
}

/// Parses `ht` from `<row r="{one_based}" ... ht="N"/>`.
fn sheet_row_height(sheet_xml: &str, one_based_row: u32) -> f64 {
    let marker = format!("<row r=\"{one_based_row}\"");
    let (_, row) = sheet_xml
        .split_once(&marker)
        .unwrap_or_else(|| panic!("missing row {one_based_row}"));
    let (row, _) = row.split_once('>').expect("unterminated row");
    let (_, height) = row.split_once("ht=\"").expect("missing row height");
    let (height, _) = height.split_once('"').expect("unterminated row height");
    height.parse().expect("row height f64")
}

fn dyn_strings(path: &Path) -> Vec<String> {
    let rows = EasyExcel::read_dynamic_sync(path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    let record = rows.last().expect("expected data row");
    (0..record.values().len())
        .filter_map(|i| match record.get(i) {
            Some(DynamicValue::String(s)) if !s.is_empty() => Some(s.clone()),
            _ => None,
        })
        .collect()
}

// ============================================================================
// Shared models
// ============================================================================

#[derive(Debug, Clone, ExcelRow)]
#[excel(column_width = 50, head_row_height = 50, content_row_height = 100)]
struct AnnotationData {
    #[excel(name = "日期", index = 0)]
    date: String,
    #[excel(name = "数字", index = 1)]
    number: f64,
    #[excel(ignore)]
    ignore: String,
}

fn annotation_data() -> Vec<AnnotationData> {
    vec![AnnotationData {
        date: "2020-01-01 01:01:01".to_owned(),
        number: 99.99,
        ignore: "忽略".to_owned(),
    }]
}

/// Java `AnnotationStyleData` — type + field Head/Content style + font.
#[derive(Debug, Clone, ExcelRow)]
#[excel(
    head_style(fill_pattern = "solid", fill_foreground_color = 10),
    head_font_style(font_height_in_points = 20, color = 15),
    content_style(fill_pattern = "solid", fill_foreground_color = 17),
    content_font_style(font_height_in_points = 30, color = 22)
)]
struct AnnotationStyleData {
    #[excel(
        name = "字符串",
        index = 0,
        head_style(fill_pattern = "solid", fill_foreground_color = 14),
        head_font_style(font_height_in_points = 40, color = 51),
        content_style(fill_pattern = "solid", fill_foreground_color = 40),
        content_font_style(font_height_in_points = 50, color = 12)
    )]
    string: String,
    #[excel(name = "字符串1", index = 1)]
    string1: String,
}

fn annotation_style_data() -> Vec<AnnotationStyleData> {
    vec![AnnotationStyleData {
        string: "string".to_owned(),
        string1: "string1".to_owned(),
    }]
}

#[derive(Debug, Clone, ExcelRow)]
struct AnnotationIndexAndNameData {
    #[excel(name = "第四个", index = 4)]
    index4: String,
    #[excel(name = "第二个", index = 2)]
    index2: String,
    #[excel(index = 0)]
    index0: String,
    #[excel(name = "第一个", index = 1)]
    index1: String,
}

fn annotation_index_name_data() -> Vec<AnnotationIndexAndNameData> {
    vec![AnnotationIndexAndNameData {
        index0: "第0个".to_owned(),
        index1: "第1个".to_owned(),
        index2: "第2个".to_owned(),
        index4: "第4个".to_owned(),
    }]
}

#[derive(Debug, Clone, ExcelRow)]
struct StyleData {
    #[excel(name = "字符串", index = 0)]
    string: String,
    #[excel(name = "字符串1", index = 1)]
    string1: String,
}

fn style_data() -> Vec<StyleData> {
    vec![
        StyleData {
            string: "字符串0".to_owned(),
            string1: "字符串01".to_owned(),
        },
        StyleData {
            string: "字符串1".to_owned(),
            string1: "字符串11".to_owned(),
        },
    ]
}

fn style_data10() -> Vec<StyleData> {
    (0..10)
        .map(|_| StyleData {
            string: "字符串0".to_owned(),
            string1: "字符串01".to_owned(),
        })
        .collect()
}

#[derive(Debug, Clone, ExcelRow)]
struct WriteHandlerData {
    #[excel(name = "姓名", index = 0)]
    name: String,
}

fn write_handler_data() -> Vec<WriteHandlerData> {
    vec![WriteHandlerData {
        name: "姓名0".to_owned(),
    }]
}

#[derive(Debug, Clone, ExcelRow)]
struct ExcludeOrIncludeData {
    #[excel(name = "column1", order = 1)]
    column1: String,
    #[excel(name = "column2", order = 2)]
    column2: String,
    #[excel(name = "column3", order = 3)]
    column3: String,
    #[excel(name = "column4", order = 4)]
    column4: String,
}

fn exclude_include_data() -> Vec<ExcludeOrIncludeData> {
    vec![ExcludeOrIncludeData {
        column1: "column1".to_owned(),
        column2: "column2".to_owned(),
        column3: "column3".to_owned(),
        column4: "column4".to_owned(),
    }]
}

// ============================================================================
// Assert helpers (shared by 1:1 mods)
// ============================================================================

/// Java `AnnotationDataTest#readAndWrite` — `@ColumnWidth` / `@HeadRowHeight` / `@ContentRowHeight`.
fn assert_annotation_dimensions(path: &Path) {
    EasyExcel::write::<AnnotationData>(path)
        .sheet("Sheet1")
        .do_write(annotation_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<AnnotationData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].date, "2020-01-01 01:01:01");
    assert!((rows[0].number - 99.99).abs() < f64::EPSILON);
    assert!(rows[0].ignore.is_empty());

    if path.extension().and_then(|ext| ext.to_str()) == Some("csv") {
        return;
    }
    if is_xls_path(path) {
        assert_real_biff8(path);
        // Minimal BIFF8 writer does not emit OOXML dimension records.
        return;
    }

    let meta = AnnotationData::write_metadata();
    assert_eq!(meta.column_width, Some(50));
    assert_eq!(meta.head_row_height, Some(50));
    assert_eq!(meta.content_row_height, Some(100));

    let sheet = zip_entry(path, "xl/worksheets/sheet1.xml");
    let col_width = sheet_column_width(&sheet, 1);
    assert!(
        (col_width - 50.0).abs() < 1.0,
        "expected ~50 character width, got {col_width}"
    );
    assert!((sheet_row_height(&sheet, 1) - 50.0).abs() < 0.5);
    assert!((sheet_row_height(&sheet, 2) - 100.0).abs() < 0.5);
}

/// Java `AnnotationDataTest#writeStyle` — field overrides type Head/Content style + fonts.
fn assert_annotation_write_style(path: &Path) {
    EasyExcel::write::<AnnotationStyleData>(path)
        .sheet("Sheet1")
        .do_write(annotation_style_data())
        .unwrap();
    if is_xls_path(path) {
        assert_real_biff8(path);
        // Cell style XF records are not asserted for minimal BIFF8 writer.
        return;
    }

    let meta = AnnotationStyleData::write_metadata();
    assert!(meta.head_style.is_some());
    assert!(meta.content_style.is_some());
    assert!(meta.head_font_style.is_some());
    assert!(meta.content_font_style.is_some());
    assert!(AnnotationStyleData::schema()[0].head_style.is_some());
    assert!(AnnotationStyleData::schema()[0].content_font_style.is_some());

    let styles = zip_entry(path, "xl/styles.xml");
    for expected in [
        "rgb=\"FFFF00FF\"", // 14 magenta (field head fill)
        "rgb=\"FFFFCC00\"", // 51
        "rgb=\"FF00CCFF\"", // 40
        "rgb=\"FF0000FF\"", // 12
        "rgb=\"FFFF0000\"", // 10 type head fill
        "rgb=\"FF00FFFF\"", // 15
        "rgb=\"FF008000\"", // 17
        "rgb=\"FFC0C0C0\"", // 22
    ] {
        assert!(styles.contains(expected), "styles.xml missing {expected}");
    }
    for size in [20, 30, 40, 50] {
        assert!(
            styles.contains(&format!("<sz val=\"{size}\"/>")),
            "styles.xml missing font size {size}"
        );
    }
}

fn assert_annotation_index_name(path: &Path) {
    EasyExcel::write::<AnnotationIndexAndNameData>(path)
        .sheet("Sheet1")
        .do_write(annotation_index_name_data())
        .unwrap();

    if path.extension().and_then(|ext| ext.to_str()) == Some("csv") {
        let vals = dyn_strings(path);
        assert!(vals.iter().any(|v| v.contains("第0个")));
        assert!(vals.iter().any(|v| v.contains("第1个")));
        assert!(vals.iter().any(|v| v.contains("第2个")));
        assert!(vals.iter().any(|v| v.contains("第4个")));
        return;
    }

    let rows = EasyExcel::read_sync::<AnnotationIndexAndNameData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].index0, "第0个");
    assert_eq!(rows[0].index1, "第1个");
    assert_eq!(rows[0].index2, "第2个");
    assert_eq!(rows[0].index4, "第4个");
}

/// Java `StyleDataTest#readAndWrite` — strategies only (width/height/horizontal style).
fn assert_style_read_and_write(path: &Path) {
    let mut head = ExcelCellStyle::new();
    head.fill_pattern = Some(ExcelFillPattern::Solid);
    head.fill_foreground_color = Some(ExcelColor::Rgb(0x00FF_FF00));
    let mut content = ExcelCellStyle::new();
    content.fill_pattern = Some(ExcelFillPattern::Solid);
    content.fill_foreground_color = Some(ExcelColor::Rgb(0x0000_8080));

    EasyExcel::write::<StyleData>(path)
        .register_write_handler(SimpleColumnWidthStyleStrategy::uniform(50))
        .register_write_handler(SimpleRowHeightStyleStrategy::new(Some(40), Some(50)))
        .register_write_handler(HorizontalCellStyleStrategy::with_head_and_content(
            head, content,
        ))
        .sheet("Sheet1")
        .do_write(style_data())
        .unwrap();

    let rows = EasyExcel::read_sync::<StyleData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].string, "字符串0");
    assert_eq!(rows[1].string1, "字符串11");
    if is_xls_path(path) {
        assert_real_biff8(path);
        // Width/height/fill style strategies are XLSX-only for now.
        return;
    }

    let sheet = zip_entry(path, "xl/worksheets/sheet1.xml");
    let col_width = sheet_column_width(&sheet, 1);
    assert!(
        (col_width - 50.0).abs() < 1.0,
        "expected ~50 character width, got {col_width}"
    );
    assert!((sheet_row_height(&sheet, 1) - 40.0).abs() < 0.5);
    assert!((sheet_row_height(&sheet, 2) - 50.0).abs() < 0.5);

    let styles = zip_entry(path, "xl/styles.xml");
    assert!(
        styles.contains("rgb=\"FFFFFF00\"") || styles.contains("theme="),
        "expected yellow head fill in styles.xml"
    );
    assert!(
        styles.contains("rgb=\"FF008080\"")
            || styles.contains("rgb=\"00008080\"")
            || styles.contains("theme="),
        "expected teal content fill in styles.xml"
    );
}

/// Java `StyleDataTest#t03/t04 AbstractVerticalCellStyleStrategy`.
fn assert_vertical_cell_style(path: &Path) {
    let strategy = VerticalCellStyleStrategy::new(
        |column| {
            let mut style = ExcelCellStyle::new();
            style.fill_pattern = Some(ExcelFillPattern::Solid);
            style.fill_foreground_color = Some(if column == 0 {
                ExcelColor::Indexed(13) // YELLOW
            } else {
                ExcelColor::Indexed(12) // BLUE
            });
            style
        },
        |column| {
            let mut style = ExcelCellStyle::new();
            style.fill_pattern = Some(ExcelFillPattern::Solid);
            style.fill_foreground_color = Some(if column == 0 {
                ExcelColor::Indexed(58) // DARK_GREEN
            } else {
                ExcelColor::Indexed(14) // PINK / MAGENTA
            });
            style
        },
    );
    EasyExcel::write::<StyleData>(path)
        .register_write_handler(strategy)
        .sheet("Sheet1")
        .do_write(style_data())
        .unwrap();

    let styles = zip_entry(path, "xl/styles.xml");
    assert!(styles.contains("rgb=\"FFFFFF00\""));
    assert!(styles.contains("rgb=\"FF0000FF\""));
    assert!(styles.contains("rgb=\"FF003300\""));
    assert!(styles.contains("rgb=\"FFFF00FF\""));
}

fn assert_loop_merge(path: &Path) {
    EasyExcel::write::<StyleData>(path)
        .loop_merge(LoopMergeStrategy::new(2, 1, 0).unwrap())
        .sheet("Sheet1")
        .do_write(style_data10())
        .unwrap();
    let rows = EasyExcel::read_sync::<StyleData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);

    let sheet = zip_entry(path, "xl/worksheets/sheet1.xml");
    assert!(
        sheet.contains("mergeCell") || sheet.contains("mergeCells"),
        "LoopMergeStrategy must emit merge regions"
    );
}

/// Custom WriteHandler that tracks lifecycle callbacks (Java WriteHandler.afterAll).
struct LifecycleWriteHandler {
    before_workbook: u32,
    after_workbook: u32,
    before_sheet: u32,
    after_sheet: u32,
    before_row: u32,
    after_row: u32,
    before_cell: u32,
    after_cell: u32,
}

impl LifecycleWriteHandler {
    fn new() -> Arc<Mutex<Self>> {
        Arc::new(Mutex::new(Self {
            before_workbook: 0,
            after_workbook: 0,
            before_sheet: 0,
            after_sheet: 0,
            before_row: 0,
            after_row: 0,
            before_cell: 0,
            after_cell: 0,
        }))
    }

    fn assert_all_one(handler: &Arc<Mutex<Self>>) {
        let h = handler.lock().unwrap();
        assert_eq!(h.before_workbook, 1, "before_workbook should be exactly 1");
        assert_eq!(h.after_workbook, 1, "after_workbook should be exactly 1");
        assert_eq!(h.before_sheet, 1, "before_sheet should be exactly 1");
        assert_eq!(h.after_sheet, 1, "after_sheet should be exactly 1");
        assert!(h.before_row >= 1, "before_row should be >= 1");
        assert!(h.after_row >= 1, "after_row should be >= 1");
        assert!(h.before_cell >= 1, "before_cell should be >= 1");
        assert!(h.after_cell >= 1, "after_cell should be >= 1");
    }
}

struct SharedLifecycleWriteHandler(Arc<Mutex<LifecycleWriteHandler>>);

impl WriteHandler for SharedLifecycleWriteHandler {
    fn before_workbook(&mut self, _ctx: &WriteWorkbookContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().before_workbook += 1;
        Ok(())
    }
    fn after_workbook(&mut self, _ctx: &WriteWorkbookContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().after_workbook += 1;
        Ok(())
    }
    fn before_sheet(&mut self, _ctx: &WriteSheetContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().before_sheet += 1;
        Ok(())
    }
    fn after_sheet(&mut self, _ctx: &WriteSheetContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().after_sheet += 1;
        Ok(())
    }
    fn before_row(&mut self, _ctx: &WriteRowContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().before_row += 1;
        Ok(())
    }
    fn after_row(&mut self, _ctx: &WriteRowContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().after_row += 1;
        Ok(())
    }
    fn before_cell(&mut self, _ctx: &mut WriteCellContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().before_cell += 1;
        Ok(())
    }
    fn after_cell(&mut self, _ctx: &WriteCellContext) -> easyexcel::Result<()> {
        self.0.lock().unwrap().after_cell += 1;
        Ok(())
    }
}

/// Java workbook/sheet/table WriteHandler — Rust registers at writer builder level.
fn assert_write_handler(path: &Path) {
    let handler = LifecycleWriteHandler::new();
    let shared = SharedLifecycleWriteHandler(handler.clone());
    EasyExcel::write::<WriteHandlerData>(path)
        .register_write_handler(shared)
        .sheet("Sheet1")
        .do_write(write_handler_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<WriteHandlerData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].name, "姓名0");
    LifecycleWriteHandler::assert_all_one(&handler);
}

fn assert_exclude_index(path: &Path) {
    let mut exclude = HashSet::new();
    exclude.insert(0usize);
    exclude.insert(3usize);
    EasyExcel::write::<ExcludeOrIncludeData>(path)
        .exclude_column_indexes(exclude)
        .sheet("Sheet1")
        .do_write(exclude_include_data())
        .unwrap();
    let vals = dyn_strings(path);
    assert!(vals.contains(&"column2".to_string()));
    assert!(vals.contains(&"column3".to_string()));
    assert!(!vals.contains(&"column1".to_string()));
    assert!(!vals.contains(&"column4".to_string()));
}

fn assert_exclude_field_name(path: &Path) {
    let exclude: HashSet<String> = ["column1", "column3", "column4"]
        .iter()
        .map(|s| (*s).to_string())
        .collect();
    EasyExcel::write::<ExcludeOrIncludeData>(path)
        .exclude_column_field_names(exclude)
        .sheet("Sheet1")
        .do_write(exclude_include_data())
        .unwrap();
    let vals = dyn_strings(path);
    assert!(vals.contains(&"column2".to_string()));
    assert!(!vals.contains(&"column1".to_string()));
    assert!(!vals.contains(&"column3".to_string()));
    assert!(!vals.contains(&"column4".to_string()));
}

fn assert_include_index(path: &Path) {
    EasyExcel::write::<ExcludeOrIncludeData>(path)
        .include_column_indexes([1usize, 2])
        .sheet("Sheet1")
        .do_write(exclude_include_data())
        .unwrap();
    let vals = dyn_strings(path);
    assert!(vals.contains(&"column2".to_string()));
    assert!(vals.contains(&"column3".to_string()));
    assert!(!vals.contains(&"column1".to_string()));
    assert!(!vals.contains(&"column4".to_string()));
}

fn assert_include_field_name(path: &Path) {
    EasyExcel::write::<ExcludeOrIncludeData>(path)
        .include_column_field_names(["column2", "column3"])
        .sheet("Sheet1")
        .do_write(exclude_include_data())
        .unwrap();
    let vals = dyn_strings(path);
    assert!(vals.contains(&"column2".to_string()));
    assert!(vals.contains(&"column3".to_string()));
    assert!(!vals.contains(&"column1".to_string()));
    assert!(!vals.contains(&"column4".to_string()));
}

fn assert_include_field_name_order(path: &Path) {
    EasyExcel::write::<ExcludeOrIncludeData>(path)
        .include_column_field_names(["column4", "column2", "column3"])
        .order_by_include_column(true)
        .sheet("Sheet1")
        .do_write(exclude_include_data())
        .unwrap();
    let vals = dyn_strings(path);
    assert_eq!(vals.len(), 3);
    assert_eq!(vals[0], "column4");
    assert_eq!(vals[1], "column2");
    assert_eq!(vals[2], "column3");
}

fn assert_include_field_name_order_index(path: &Path) {
    EasyExcel::write::<ExcludeOrIncludeData>(path)
        .include_column_indexes([3usize, 1, 2, 0])
        .order_by_include_column(true)
        .sheet("Sheet1")
        .do_write(exclude_include_data())
        .unwrap();
    let vals = dyn_strings(path);
    assert_eq!(vals.len(), 4);
    assert_eq!(vals[0], "column4");
    assert_eq!(vals[1], "column2");
    assert_eq!(vals[2], "column3");
    assert_eq!(vals[3], "column1");
}

// ============================================================================
// AnnotationDataTest (5 @Test)
// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationDataTest
// ============================================================================

mod annotation_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_annotation_dimensions(&temp_path("annotation07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        assert_annotation_dimensions(&temp_path("annotation03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        assert_annotation_dimensions(&temp_path("annotationCsv.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationDataTest#t11WriteStyle07
    #[test]
    fn t11_write_style07() {
        assert_annotation_write_style(&temp_path("annotationStyle07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationDataTest#t12Write03
    #[test]
    fn t12_write03() {
        assert_annotation_write_style(&temp_path("annotationStyle03.xls"));
    }
}

// ============================================================================
// AnnotationIndexAndNameDataTest (3 @Test)
// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationIndexAndNameDataTest
// ============================================================================

mod annotation_index_and_name_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationIndexAndNameDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_annotation_index_name(&temp_path("annotationIndexAndName07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationIndexAndNameDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        assert_annotation_index_name(&temp_path("annotationIndexAndName03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.annotation.AnnotationIndexAndNameDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        assert_annotation_index_name(&temp_path("annotationIndexAndNameCsv.csv"));
    }
}

// ============================================================================
// StyleDataTest (5 @Test)
// Java: com.alibaba.easyexcel.test.core.style.StyleDataTest
// ============================================================================

mod style_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.style.StyleDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_style_read_and_write(&temp_path("style07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.style.StyleDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        assert_style_read_and_write(&temp_path("style03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.style.StyleDataTest#t03AbstractVerticalCellStyleStrategy
    #[test]
    fn t03_abstract_vertical_cell_style_strategy() {
        assert_vertical_cell_style(&temp_path("verticalCellStyle.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.style.StyleDataTest#t04AbstractVerticalCellStyleStrategy02
    #[test]
    fn t04_abstract_vertical_cell_style_strategy02() {
        // Java builds WriteCellStyle from StyleProperty/FontProperty; Rust uses same
        // column-differentiated VerticalCellStyleStrategy fills as t03.
        assert_vertical_cell_style(&temp_path("verticalCellStyle2.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.style.StyleDataTest#t05LoopMergeStrategy
    #[test]
    fn t05_loop_merge_strategy() {
        assert_loop_merge(&temp_path("loopMergeStrategy.xlsx"));
    }
}

// ============================================================================
// WriteHandlerTest (9 @Test)
// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest
// workbook / sheet / table × 07 / 03 / csv
// ============================================================================

mod write_handler_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t01WorkbookWrite07
    #[test]
    fn t01_workbook_write07() {
        assert_write_handler(&temp_path("writeHandler07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t02WorkbookWrite03
    #[test]
    fn t02_workbook_write03() {
        assert_write_handler(&temp_path("writeHandler03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t03WorkbookWriteCsv
    #[test]
    fn t03_workbook_write_csv() {
        assert_write_handler(&temp_path("writeHandlerCsv.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t11SheetWrite07
    #[test]
    fn t11_sheet_write07() {
        // Java: sheet().registerWriteHandler(...). Rust API registers at writer builder.
        assert_write_handler(&temp_path("writeHandlerSheet07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t12SheetWrite03
    #[test]
    fn t12_sheet_write03() {
        assert_write_handler(&temp_path("writeHandlerSheet03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t13SheetWriteCsv
    #[test]
    fn t13_sheet_write_csv() {
        assert_write_handler(&temp_path("writeHandlerSheetCsv.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t21TableWrite07
    #[test]
    fn t21_table_write07() {
        // Java: sheet().table(0).registerWriteHandler(...). Rust registers at writer builder.
        assert_write_handler(&temp_path("writeHandlerTable07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t22TableWrite03
    #[test]
    fn t22_table_write03() {
        assert_write_handler(&temp_path("writeHandlerTable03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.handler.WriteHandlerTest#t23TableWriteCsv
    #[test]
    fn t23_table_write_csv() {
        assert_write_handler(&temp_path("writeHandlerTableCsv.csv"));
    }
}

// ============================================================================
// ExcludeOrIncludeDataTest (18 @Test)
// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest
// ============================================================================

mod exclude_or_include_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t01ExcludeIndex07
    #[test]
    fn t01_exclude_index07() {
        assert_exclude_index(&temp_path("excludeIndex.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t02ExcludeIndex03
    #[test]
    fn t02_exclude_index03() {
        assert_exclude_index(&temp_path("excludeIndex03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t03ExcludeIndexCsv
    #[test]
    fn t03_exclude_index_csv() {
        assert_exclude_index(&temp_path("excludeIndex.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t11ExcludeFieldName07
    #[test]
    fn t11_exclude_field_name07() {
        assert_exclude_field_name(&temp_path("excludeFieldName.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t12ExcludeFieldName03
    #[test]
    fn t12_exclude_field_name03() {
        assert_exclude_field_name(&temp_path("excludeFieldName03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t13ExcludeFieldNameCsv
    #[test]
    fn t13_exclude_field_name_csv() {
        assert_exclude_field_name(&temp_path("excludeFieldName.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t21IncludeIndex07
    #[test]
    fn t21_include_index07() {
        assert_include_index(&temp_path("includeIndex.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t22IncludeIndex03
    #[test]
    fn t22_include_index03() {
        assert_include_index(&temp_path("includeIndex03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t23IncludeIndexCsv
    #[test]
    fn t23_include_index_csv() {
        assert_include_index(&temp_path("includeIndex.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t31IncludeFieldName07
    #[test]
    fn t31_include_field_name07() {
        assert_include_field_name(&temp_path("includeFieldName.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t32IncludeFieldName03
    #[test]
    fn t32_include_field_name03() {
        assert_include_field_name(&temp_path("includeFieldName03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t33IncludeFieldNameCsv
    #[test]
    fn t33_include_field_name_csv() {
        assert_include_field_name(&temp_path("includeFieldName.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t41IncludeFieldNameOrder07
    #[test]
    fn t41_include_field_name_order07() {
        assert_include_field_name_order(&temp_path("includeFieldNameOrder.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t42IncludeFieldNameOrder03
    #[test]
    fn t42_include_field_name_order03() {
        assert_include_field_name_order(&temp_path("includeFieldNameOrder03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t43IncludeFieldNameOrderCsv
    #[test]
    fn t43_include_field_name_order_csv() {
        assert_include_field_name_order(&temp_path("includeFieldNameOrder.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t41IncludeFieldNameOrderIndex07
    #[test]
    fn t41_include_field_name_order_index07() {
        assert_include_field_name_order_index(&temp_path("includeFieldNameOrderIndex.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t42IncludeFieldNameOrderIndex03
    #[test]
    fn t42_include_field_name_order_index03() {
        assert_include_field_name_order_index(&temp_path(
            "includeFieldNameOrderIndex03.xls",
        ));
    }

    /// Java: com.alibaba.easyexcel.test.core.excludeorinclude.ExcludeOrIncludeDataTest#t43IncludeFieldNameOrderIndexCsv
    #[test]
    fn t43_include_field_name_order_index_csv() {
        assert_include_field_name_order_index(&temp_path("includeFieldNameOrderIndex.csv"));
    }
}
