use std::collections::HashMap;
use std::fs;
use std::io::{BufRead, Read, Write};
use std::sync::Arc;

use base64::Engine;
use calamine::{CellErrorType, ExcelDateTime, ExcelDateTimeType};
use easyexcel_core::{DynamicRow, DynamicValue, ExcelColumn, IntoExcelCell};
use flate2::read::GzDecoder;
use rust_xlsxwriter::{Format, Note, Workbook};
use tempfile::{TempDir, tempdir};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use super::*;

struct FaultyBufRead;

impl Read for FaultyBufRead {
    fn read(&mut self, _buffer: &mut [u8]) -> std::io::Result<usize> {
        Err(std::io::Error::other("injected probe failure"))
    }
}

impl BufRead for FaultyBufRead {
    fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
        Err(std::io::Error::other("injected probe failure"))
    }

    fn consume(&mut self, _amount: usize) {}
}

fn test_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[derive(Debug, PartialEq, Eq)]
struct TestRow(String);

impl ExcelRow for TestRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Value", Some(0), 0, None)];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        let value = row
            .cell(&Self::schema()[0])
            .map_or_else(String::new, CellValue::as_text);
        if value == "conversion-error" {
            return Err(ExcelError::Format("conversion failed".to_owned()));
        }
        Ok(Self(value))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        self.0
            .to_excel_cell(&easyexcel_core::ConvertContext {
                sheet_name: String::new(),
                row_index: 0,
                column_index: Some(0),
                field: "value",
                format: None,
            })
            .map(|value| vec![value])
    }
}

#[derive(Debug, PartialEq, Eq)]
struct NamedRow(String);

impl ExcelRow for NamedRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Canonical", None, 0, None)];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        Ok(Self(
            row.cell(&Self::schema()[0])
                .map_or_else(String::new, CellValue::as_text),
        ))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![CellValue::String(self.0.clone())])
    }
}

#[derive(Default)]
struct NamedProbe {
    heads: Vec<HashMap<String, usize>>,
    rows: Vec<NamedRow>,
}

impl ReadListener<NamedRow> for NamedProbe {
    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        self.heads.push(head.clone());
        Ok(())
    }

    fn invoke(&mut self, data: NamedRow, _context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq)]
struct RawRow {
    cells: Vec<CellValue>,
    formulas: Vec<Option<String>>,
}

impl ExcelRow for RawRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[
            ExcelColumn::new("shared", "Shared", Some(0), 0, None),
            ExcelColumn::new("inline", "Inline", Some(1), 1, None),
            ExcelColumn::new("boolean", "Boolean", Some(2), 2, None),
            ExcelColumn::new("integer", "Integer", Some(3), 3, None),
            ExcelColumn::new("float", "Float", Some(4), 4, None),
            ExcelColumn::new("formula_number", "Formula number", Some(5), 5, None),
            ExcelColumn::new("formula_string", "Formula string", Some(6), 6, None),
            ExcelColumn::new("error", "Error", Some(7), 7, None),
            ExcelColumn::new("date", "Date", Some(8), 8, None),
        ];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        Ok(Self {
            cells: Self::schema()
                .iter()
                .map(|column| row.cell(column).cloned().unwrap_or(CellValue::Empty))
                .collect(),
            formulas: Self::schema()
                .iter()
                .map(|column| {
                    row.formula(column)
                        .map(|formula| formula.formula_value().to_owned())
                })
                .collect(),
        })
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(self.cells.clone())
    }
}

#[derive(Default)]
struct RawProbe(Vec<RawRow>);

impl ReadListener<RawRow> for RawProbe {
    fn invoke(&mut self, data: RawRow, _context: &AnalysisContext) -> Result<()> {
        self.0.push(data);
        Ok(())
    }
}

#[derive(Default)]
struct DynamicProbe(Vec<DynamicRow>);

impl ReadListener<DynamicRow> for DynamicProbe {
    fn invoke(&mut self, data: DynamicRow, _context: &AnalysisContext) -> Result<()> {
        self.0.push(data);
        Ok(())
    }
}

#[derive(Default)]
#[allow(clippy::struct_excessive_bools)]
struct Probe {
    heads: Vec<HashMap<String, usize>>,
    rows: Vec<TestRow>,
    after: Vec<(String, usize, u32)>,
    continue_reading: bool,
    fail_head: bool,
    fail_invoke: bool,
    fail_invoke_at: Option<usize>,
    invoke_count: usize,
    fail_after: bool,
    error_action: Option<ErrorAction>,
    errors: usize,
    stop_after_callbacks: Option<usize>,
    callback_count: usize,
}

impl ReadListener<TestRow> for Probe {
    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        if self.fail_head {
            return Err(ExcelError::Format("head failed".to_owned()));
        }
        self.heads.push(head.clone());
        Ok(())
    }

    fn invoke(&mut self, data: TestRow, _context: &AnalysisContext) -> Result<()> {
        self.invoke_count += 1;
        if self.fail_invoke || self.fail_invoke_at == Some(self.invoke_count) {
            return Err(ExcelError::Format("invoke failed".to_owned()));
        }
        self.rows.push(data);
        Ok(())
    }

    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        self.errors += 1;
        self.error_action.unwrap_or(ErrorAction::Stop)
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        if self.fail_after {
            return Err(ExcelError::Format("after failed".to_owned()));
        }
        self.after.push((
            context.sheet_name().to_owned(),
            context.sheet_no(),
            context.row_index(),
        ));
        Ok(())
    }

    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        self.callback_count += 1;
        self.stop_after_callbacks
            .map_or(self.continue_reading, |limit| self.callback_count < limit)
    }
}

#[derive(Default)]
struct ExtraProbe {
    events: Vec<&'static str>,
    extras: Vec<CellExtra>,
    context_customs: Vec<Option<String>>,
    fail_extra: bool,
    error_action: Option<ErrorAction>,
    errors: usize,
    stop_after_extra: bool,
    extra_seen: bool,
}

impl ExtraProbe {
    fn record_custom(&mut self, context: &AnalysisContext) {
        self.context_customs
            .push(context.custom::<String>().cloned());
    }
}

impl ReadListener<TestRow> for ExtraProbe {
    fn on_exception(&mut self, _error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        self.record_custom(context);
        self.errors += 1;
        self.error_action.unwrap_or(ErrorAction::Stop)
    }

    fn invoke_head(
        &mut self,
        _head: &HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        self.record_custom(context);
        self.events.push("head");
        Ok(())
    }

    fn invoke(&mut self, _data: TestRow, context: &AnalysisContext) -> Result<()> {
        self.record_custom(context);
        self.events.push("row");
        Ok(())
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        self.record_custom(context);
        self.events.push("extra");
        self.extras.push(extra.clone());
        self.extra_seen = true;
        if self.fail_extra {
            Err(ExcelError::Format("extra failed".to_owned()))
        } else {
            Ok(())
        }
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.record_custom(context);
        self.events.push("after");
        Ok(())
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        self.record_custom(context);
        !(self.stop_after_extra && self.extra_seen)
    }
}

struct ErrorProbe {
    action: ErrorAction,
    errors: usize,
}

impl ReadListener<TestRow> for ErrorProbe {
    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        self.errors += 1;
        self.action
    }

    fn invoke(&mut self, _data: TestRow, _context: &AnalysisContext) -> Result<()> {
        panic!("a conversion failure cannot invoke a row")
    }
}

fn options() -> ReadOptions {
    ReadOptions {
        sheet: SheetSelector::First,
        head_row_number: 1,
        ignore_empty_row: true,
        auto_trim: true,
        use_1904_windowing: false,
        scientific_format: ScientificFormatMode::Plain,
        locale: ExcelLocale::default(),
        start_row: None,
        end_row: None,
        header_aliases: HashMap::new(),
        custom_object: None,
        read_default_return: ReadDefaultReturn::default(),
        extra_read: HashSet::new(),
        password: None,
        charset: CsvCharset::default(),
        converters: easyexcel_core::ConverterRegistry::default(),
        read_cache: ReadCacheMode::default(),
    }
}

fn workbook_fixture() -> Result<(TempDir, std::path::PathBuf)> {
    let directory = tempdir()?;
    let path = directory.path().join("fixture.xlsx");
    let mut workbook = Workbook::new();
    let first = workbook.add_worksheet();
    first.set_name("First").map_err(test_error)?;
    first.write_string(0, 0, "Value").map_err(test_error)?;
    first.write_string(1, 0, "one").map_err(test_error)?;
    let second = workbook.add_worksheet();
    second.set_name("Second").map_err(test_error)?;
    second.write_string(0, 0, "Value").map_err(test_error)?;
    second.write_string(1, 0, "two").map_err(test_error)?;
    workbook.save(&path).map_err(test_error)?;
    Ok((directory, path))
}

fn extra_workbook_fixture() -> Result<(TempDir, std::path::PathBuf)> {
    let directory = tempdir()?;
    let path = directory.path().join("extras.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("Meta").map_err(test_error)?;
    worksheet.write_string(0, 0, "Value").map_err(test_error)?;
    worksheet.write_string(1, 0, "row").map_err(test_error)?;
    worksheet
        .insert_note(1, 0, &Note::new("comment & text"))
        .map_err(test_error)?;
    worksheet
        .write_url(2, 0, "https://example.com")
        .map_err(test_error)?;
    worksheet
        .write_url(2, 1, "internal:Meta!A1")
        .map_err(test_error)?;
    worksheet
        .merge_range(3, 0, 3, 1, "Merged", &Format::new())
        .map_err(test_error)?;
    workbook.save(&path).map_err(test_error)?;
    Ok((directory, path))
}

fn rewrite_first_sheet(source: &Path, destination: &Path, replacement: &str) -> Result<()> {
    let mut archive = ZipArchive::new(fs::File::open(source)?).map_err(test_error)?;
    let mut writer = ZipWriter::new(fs::File::create(destination)?);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).map_err(test_error)?;
        let name = entry.name().to_owned();
        let options = SimpleFileOptions::default().compression_method(entry.compression());
        if entry.is_dir() {
            writer.add_directory(name, options).map_err(test_error)?;
            continue;
        }
        writer.start_file(&name, options).map_err(test_error)?;
        if name == "xl/worksheets/sheet1.xml" {
            writer
                .write_all(replacement.as_bytes())
                .map_err(test_error)?;
        } else {
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            writer.write_all(&bytes)?;
        }
    }
    writer.finish().map_err(test_error)?;
    Ok(())
}

fn write_xlsx_package(path: &Path, entries: &[(&str, &str)]) -> Result<()> {
    let mut writer = ZipWriter::new(fs::File::create(path)?);
    for (name, contents) in entries {
        writer
            .start_file(*name, SimpleFileOptions::default())
            .map_err(test_error)?;
        writer.write_all(contents.as_bytes())?;
    }
    writer.finish().map_err(test_error)?;
    Ok(())
}

fn remove_first_sheet(source: &Path, destination: &Path) -> Result<()> {
    let mut archive = ZipArchive::new(fs::File::open(source)?).map_err(test_error)?;
    let mut writer = ZipWriter::new(fs::File::create(destination)?);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).map_err(test_error)?;
        let name = entry.name().to_owned();
        if name == "xl/worksheets/sheet1.xml" {
            continue;
        }
        let options = SimpleFileOptions::default().compression_method(entry.compression());
        if entry.is_dir() {
            writer.add_directory(name, options).map_err(test_error)?;
            continue;
        }
        writer.start_file(&name, options).map_err(test_error)?;
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        writer.write_all(&bytes)?;
    }
    writer.finish().map_err(test_error)?;
    Ok(())
}

fn worksheet_xml(cells: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1">{cells}</row></sheetData>
</worksheet>"#
    )
}

fn column_name(index: u32) -> String {
    let mut value = index + 1;
    let mut name = String::new();
    while value > 0 {
        let remainder = ((value - 1) % 26) as u8;
        name.insert(0, char::from(b'A' + remainder));
        value = (value - 1) / 26;
    }
    name
}

fn encode_csv_fixture(encoding: &'static encoding_rs::Encoding, value: &str) -> Vec<u8> {
    if encoding == encoding_rs::UTF_16BE {
        value.encode_utf16().flat_map(u16::to_be_bytes).collect()
    } else if encoding == encoding_rs::UTF_16LE {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    } else {
        let (encoded, actual, had_errors) = encoding.encode(value);
        assert_eq!(actual, encoding);
        assert!(!had_errors);
        encoded.into_owned()
    }
}

#[test]
fn calamine_values_map_to_every_core_cell_variant() {
    let datetime = ExcelDateTime::new(46_120.5, ExcelDateTimeType::DateTime, false);
    let invalid_datetime = ExcelDateTime::new(f64::MAX, ExcelDateTimeType::DateTime, false);
    let duration = ExcelDateTime::new(1.5, ExcelDateTimeType::TimeDelta, false);
    let cases = [
        (DataRef::Empty, CellValue::Empty),
        (
            DataRef::String("owned".to_owned()),
            CellValue::String("owned".to_owned()),
        ),
        (
            DataRef::SharedString("shared"),
            CellValue::String("shared".to_owned()),
        ),
        (
            DataRef::DateTimeIso("2026-01-01".to_owned()),
            CellValue::String("2026-01-01".to_owned()),
        ),
        (
            DataRef::DurationIso("PT1H".to_owned()),
            CellValue::String("PT1H".to_owned()),
        ),
        (DataRef::Bool(true), CellValue::Bool(true)),
        (DataRef::Int(7), CellValue::Int(7)),
        (DataRef::Float(1.25), CellValue::Float(1.25)),
        (DataRef::DateTime(duration), CellValue::Float(1.5)),
        (
            DataRef::DateTime(invalid_datetime),
            CellValue::Float(f64::MAX),
        ),
        (
            DataRef::Error(CellErrorType::Div0),
            CellValue::String("#DIV/0!".to_owned()),
        ),
    ];
    for (input, expected) in cases {
        assert_eq!(from_calamine(&input, false), expected);
        assert_eq!(from_data(&Data::from(input), false), expected);
    }
    assert!(matches!(
        from_calamine(&DataRef::DateTime(datetime), false),
        CellValue::DateTime(_)
    ));
    assert!(matches!(
        from_data(&Data::DateTime(datetime), false),
        CellValue::DateTime(_)
    ));
    let serial_one = ExcelDateTime::new(1.0, ExcelDateTimeType::DateTime, true);
    assert_eq!(
        from_calamine(&DataRef::DateTime(serial_one), false).as_text(),
        "1900-01-01 00:00:00"
    );
    assert_eq!(
        from_data(&Data::DateTime(serial_one), true).as_text(),
        "1904-01-02 00:00:00"
    );
}

#[test]
fn helpers_preserve_diagnostics_and_xlsx_column_limits() {
    assert_eq!(ReadOptions::default(), options());
    assert_eq!(SheetSelector::default(), SheetSelector::First);
    assert_eq!(to_column_index(0).expect("column"), 0);
    assert_eq!(
        to_column_index(u32::from(u16::MAX)).expect("column"),
        usize::from(u16::MAX)
    );
    assert!(to_column_index(u32::from(u16::MAX) + 1).is_err());
    assert_eq!(
        format_error("broken").to_string(),
        "excel format error: broken"
    );
    assert!(!is_compound_document(&mut FaultyBufRead));
    assert_eq!(java_trim("\0\t value \r\n"), "value");
    assert_eq!(java_trim("\u{a0}value\u{a0}"), "\u{a0}value\u{a0}");
    assert!(sheet_name_matches(" Sheet ", "Sheet", true));
    assert!(!sheet_name_matches(" Sheet ", "Sheet", false));
}

#[test]
fn header_aliases_and_inclusive_row_ranges_apply_before_typed_mapping() -> Result<()> {
    let mut range = Range::new((0, 0), (3, 0));
    range.set_value((0, 0), Data::String("Source".to_owned()));
    range.set_value((1, 0), Data::String("one".to_owned()));
    range.set_value((2, 0), Data::String("two".to_owned()));
    range.set_value((3, 0), Data::String("three".to_owned()));
    let mut aliases = HashMap::new();
    aliases.insert("Source".to_owned(), "Canonical".to_owned());
    let options = ReadOptions {
        start_row: Some(2),
        end_row: Some(2),
        header_aliases: aliases,
        ..ReadOptions::default()
    };
    let mut probe = NamedProbe::default();

    assert_eq!(
        read_range(
            &range,
            0,
            "Aliased",
            &options,
            &mut TypedRowConsumer::<NamedRow> {
                listener: &mut probe,
            },
        )?,
        ReadFlow::Continue
    );
    assert_eq!(probe.heads[0].get("Canonical"), Some(&0));
    assert_eq!(probe.rows, vec![NamedRow("two".to_owned())]);
    Ok(())
}

#[test]
fn read_row_range_validation_rejects_only_reversed_bounds() {
    assert!(validate_read_options(&ReadOptions::default()).is_ok());
    assert!(
        validate_read_options(&ReadOptions {
            start_row: Some(2),
            ..ReadOptions::default()
        })
        .is_ok()
    );
    assert!(
        validate_read_options(&ReadOptions {
            end_row: Some(2),
            ..ReadOptions::default()
        })
        .is_ok()
    );
    assert!(
        validate_read_options(&ReadOptions {
            start_row: Some(2),
            end_row: Some(2),
            ..ReadOptions::default()
        })
        .is_ok()
    );
    assert_eq!(
        validate_read_options(&ReadOptions {
            start_row: Some(3),
            end_row: Some(2),
            ..ReadOptions::default()
        })
        .expect_err("reversed row range")
        .to_string(),
        "excel format error: read row range start 3 exceeds end 2"
    );

    let reversed = ReadOptions {
        start_row: Some(3),
        end_row: Some(2),
        ..ReadOptions::default()
    };
    let mut modern_workbook_probe = Probe::default();
    let mut legacy_workbook_probe = Probe::default();
    let mut delimited_text_probe = Probe::default();
    assert!(
        read_xlsx::<TestRow, _>(
            Path::new("missing.xlsx"),
            &reversed,
            &mut modern_workbook_probe,
        )
        .is_err()
    );
    assert!(
        read_xls::<TestRow, _>(
            Path::new("missing.xls"),
            &reversed,
            &mut legacy_workbook_probe,
        )
        .is_err()
    );
    assert!(
        read_csv::<TestRow, _>(
            Path::new("missing.csv"),
            &reversed,
            &mut delimited_text_probe,
        )
        .is_err()
    );
}

#[test]
#[allow(clippy::too_many_lines)]
fn legacy_range_read_preserves_coordinates_headers_and_empty_sheets() -> Result<()> {
    let mut range = Range::new((2, 1), (3, 1));
    range.set_value((2, 1), Data::String("Value".to_owned()));
    range.set_value((3, 1), Data::String("one".to_owned()));
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_range(
        &range,
        2,
        "Legacy",
        &ReadOptions {
            head_row_number: 3,
            ..options()
        },
        &mut TypedRowConsumer::<TestRow> {
            listener: &mut probe,
        },
    )?;
    assert_eq!(probe.heads[0].get("Value"), Some(&1));
    assert_eq!(probe.rows, vec![TestRow(String::new())]);
    assert_eq!(probe.after, vec![("Legacy".to_owned(), 2, 3)]);

    let mut stopped = Probe::default();
    assert_eq!(
        read_range(
            &range,
            2,
            "Legacy",
            &ReadOptions {
                head_row_number: 3,
                ..options()
            },
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut stopped,
            },
        )?,
        ReadFlow::Stop
    );
    assert_eq!(stopped.heads.len(), 1);
    assert!(stopped.rows.is_empty());
    assert!(stopped.after.is_empty());

    read_range(
        &Range::empty(),
        3,
        "Empty",
        &options(),
        &mut TypedRowConsumer::<TestRow> {
            listener: &mut probe,
        },
    )?;
    assert_eq!(probe.after.last(), Some(&("Empty".to_owned(), 3, 0)));

    let mut failing_empty_after = Probe {
        continue_reading: true,
        fail_after: true,
        ..Probe::default()
    };
    assert!(
        read_range(
            &Range::empty(),
            3,
            "Empty",
            &options(),
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut failing_empty_after,
            },
        )
        .is_err()
    );

    let mut failing_range_after = Probe {
        continue_reading: true,
        fail_after: true,
        ..Probe::default()
    };
    assert!(
        read_range(
            &range,
            2,
            "Legacy",
            &ReadOptions {
                head_row_number: 3,
                ..options()
            },
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut failing_range_after,
            },
        )
        .is_err()
    );

    let invalid_column = Range::new((0, u32::from(u16::MAX) + 1), (0, u32::from(u16::MAX) + 1));
    assert!(
        read_range(
            &invalid_column,
            0,
            "Invalid",
            &options(),
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut probe,
            },
        )
        .is_err()
    );

    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_range(
            &range,
            0,
            "Legacy",
            &ReadOptions {
                head_row_number: 3,
                ..options()
            },
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut failing_head,
            },
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn reads_java_easyexcel_legacy_multisheet_fixture() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("java-multiplesheets.xls");
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(include_str!("fixtures/java-multiplesheets.xls.gz.b64").trim())
        .map_err(test_error)?;
    let mut decoder = GzDecoder::new(compressed.as_slice());
    let mut workbook = Vec::new();
    decoder.read_to_end(&mut workbook)?;
    fs::write(&path, workbook)?;
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xls::<TestRow, _>(
        &path,
        &ReadOptions {
            sheet: SheetSelector::All,
            ..options()
        },
        &mut probe,
    )?;
    assert_eq!(
        probe.rows,
        (1..=6)
            .map(|index| TestRow(format!("表{index}数据")))
            .collect::<Vec<_>>()
    );
    assert_eq!(
        probe.after,
        (0..6)
            .map(|index| (format!("Sheet{}", index + 1), index, 1))
            .collect::<Vec<_>>()
    );
    assert_eq!(probe.heads.len(), 6);
    for (index, head) in probe.heads.iter().enumerate() {
        assert_eq!(head.get(&format!("表{}头", index + 1)), Some(&0));
    }
    let mut stopped = Probe::default();
    read_xls::<TestRow, _>(
        &path,
        &ReadOptions {
            sheet: SheetSelector::All,
            ..options()
        },
        &mut stopped,
    )?;
    assert_eq!(stopped.heads.len(), 1);
    assert!(stopped.rows.is_empty());
    assert!(stopped.after.is_empty());
    assert!(
        read_xls::<TestRow, _>(
            &path,
            &ReadOptions {
                sheet: SheetSelector::Index(99),
                ..options()
            },
            &mut probe,
        )
        .is_err()
    );
    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(read_xls::<TestRow, _>(&path, &options(), &mut failing_head).is_err());
    let invalid = directory.path().join("invalid.xls");
    fs::write(&invalid, b"not an XLS workbook")?;
    assert!(read_xls::<TestRow, _>(&invalid, &options(), &mut probe).is_err());

    let mut dynamic = DynamicProbe::default();
    read_xls::<DynamicRow, _>(
        &path,
        &ReadOptions {
            read_default_return: ReadDefaultReturn::ReadCellData,
            ..options()
        },
        &mut dynamic,
    )?;
    let DynamicValue::ReadCellData(cell) = dynamic.0[0].get(0).expect("legacy cell") else {
        panic!("expected legacy read cell data");
    };
    assert_eq!(cell.raw_value(), &CellValue::String("表1数据".to_owned()));
    assert_eq!(cell.data(), &CellValue::String("表1数据".to_owned()));
    assert_eq!(cell.row_index(), 1);
    assert_eq!(cell.column_index(), 0);
    Ok(())
}

fn java_compatibility_fixture(directory: &TempDir, name: &str) -> Result<std::path::PathBuf> {
    let encoded = match name {
        "t01.xls" => include_str!("fixtures/java-compat-t01.xls.gz.b64"),
        "t02.xlsx" => include_str!("fixtures/java-compat-t02.xlsx.gz.b64"),
        "t03.xlsx" => include_str!("fixtures/java-compat-t03.xlsx.gz.b64"),
        "t04.xlsx" => include_str!("fixtures/java-compat-t04.xlsx.gz.b64"),
        "t05.xlsx" => include_str!("fixtures/java-compat-t05.xlsx.gz.b64"),
        "t06.xlsx" => include_str!("fixtures/java-compat-t06.xlsx.gz.b64"),
        "t07.xlsx" => include_str!("fixtures/java-compat-t07.xlsx.gz.b64"),
        "t09.xlsx" => include_str!("fixtures/java-compat-t09.xlsx.gz.b64"),
        _ => {
            return Err(ExcelError::Format(format!(
                "unknown compatibility fixture: {name}"
            )));
        }
    };
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(encoded.trim())
        .map_err(test_error)?;
    let mut decoder = GzDecoder::new(compressed.as_slice());
    let mut workbook = Vec::new();
    decoder.read_to_end(&mut workbook).map_err(test_error)?;
    let path = directory.path().join(name);
    fs::write(&path, workbook).map_err(test_error)?;
    Ok(path)
}

fn read_java_compatibility_rows(
    directory: &TempDir,
    name: &str,
    head_row_number: u32,
    read_default_return: ReadDefaultReturn,
) -> Result<Vec<DynamicRow>> {
    let path = java_compatibility_fixture(directory, name)?;
    let options = ReadOptions {
        head_row_number,
        read_default_return,
        ..ReadOptions::default()
    };
    let mut listener = DynamicProbe::default();
    if path.extension().is_some_and(|extension| extension == "xls") {
        read_xls::<DynamicRow, _>(&path, &options, &mut listener)?;
    } else {
        read_xlsx::<DynamicRow, _>(&path, &options, &mut listener)?;
    }
    Ok(listener.0)
}

#[test]
fn reads_java_official_compatibility_fixtures() -> Result<()> {
    let directory = tempdir().map_err(test_error)?;
    let t01 = read_java_compatibility_rows(&directory, "t01.xls", 1, ReadDefaultReturn::String)?;
    assert_eq!(t01.len(), 2);
    assert_eq!(
        t01[1].get(0),
        Some(&DynamicValue::String("Q235(碳钢)".to_owned()))
    );

    let t02 = read_java_compatibility_rows(&directory, "t02.xlsx", 0, ReadDefaultReturn::String)?;
    assert_eq!(t02.len(), 3);
    assert_eq!(
        t02[2].get(2),
        Some(&DynamicValue::String("1，2-戊二醇".to_owned()))
    );

    let t03 = read_java_compatibility_rows(&directory, "t03.xlsx", 1, ReadDefaultReturn::String)?;
    assert_eq!(t03.len(), 1);
    assert_eq!(t03[0].values().len(), 12);

    let t04 = read_java_compatibility_rows(&directory, "t04.xlsx", 1, ReadDefaultReturn::String)?;
    assert_eq!(t04.len(), 56);
    assert_eq!(
        t04[0].get(5),
        Some(&DynamicValue::String("QQSJK28F152A012242S0081".to_owned()))
    );

    let t05 = read_java_compatibility_rows(&directory, "t05.xlsx", 1, ReadDefaultReturn::String)?;
    for (row, expected) in [
        "2023-01-01 00:00:00",
        "2023-01-01 00:00:00",
        "2023-01-01 00:00:00",
        "2023-01-01 00:00:01",
        "2023-01-01 00:00:01",
    ]
    .into_iter()
    .enumerate()
    {
        assert_eq!(
            t05[row].get(0),
            Some(&DynamicValue::String(expected.to_owned()))
        );
    }

    let t06 = read_java_compatibility_rows(&directory, "t06.xlsx", 0, ReadDefaultReturn::String)?;
    assert_eq!(
        t06[0].get(2),
        Some(&DynamicValue::String("2087.03".to_owned()))
    );

    let t07_actual =
        read_java_compatibility_rows(&directory, "t07.xlsx", 1, ReadDefaultReturn::ActualData)?;
    let Some(DynamicValue::ActualData(CellValue::Decimal(actual))) = t07_actual[0].get(11) else {
        panic!("expected actual decimal value");
    };
    assert_eq!(actual.to_string(), "24.1998124");
    let t07_string =
        read_java_compatibility_rows(&directory, "t07.xlsx", 1, ReadDefaultReturn::String)?;
    assert_eq!(
        t07_string[0].get(11),
        Some(&DynamicValue::String("24.20".to_owned()))
    );

    let t09 = read_java_compatibility_rows(&directory, "t09.xlsx", 0, ReadDefaultReturn::String)?;
    assert_eq!(t09.len(), 1);
    assert_eq!(
        t09[0].get(0),
        Some(&DynamicValue::String("SH_x000D_Z002".to_owned()))
    );
    Ok(())
}

#[test]
fn reads_java_easyexcel_encrypted_xlsx_fixture() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("java-encrypt07.xlsx");
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(include_str!("fixtures/java-encrypt07.xlsx.gz.b64").trim())
        .map_err(test_error)?;
    let mut decoder = GzDecoder::new(compressed.as_slice());
    let mut workbook = Vec::new();
    decoder.read_to_end(&mut workbook)?;
    assert!(is_compound_document(&mut workbook.as_slice()));
    assert!(!is_compound_document(&mut &workbook[..4]));
    assert!(!is_compound_document(&mut &b"not-cfb!"[..]));
    fs::write(&path, workbook)?;

    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            password: Some("123456".to_owned()),
            extra_read: HashSet::from([CellExtraType::Merge]),
            ..options()
        },
        &mut probe,
    )?;
    assert!(read_xlsx::<TestRow, _>(&path, &options(), &mut probe).is_err());
    assert_eq!(
        probe.rows,
        (0..10)
            .map(|index| TestRow(format!("姓名{index}")))
            .collect::<Vec<_>>()
    );
    assert_eq!(probe.heads[0].get("姓名"), Some(&0));
    assert_eq!(probe.after, vec![("0".to_owned(), 0, 10)]);

    let mut empty_row_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            ignore_empty_row: false,
            password: Some("123456".to_owned()),
            ..options()
        },
        &mut empty_row_probe,
    )?;
    assert_eq!(empty_row_probe.rows.len(), 10);
    assert_eq!(empty_row_probe.after, vec![("0".to_owned(), 0, 10)]);
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn xlsx_extra_callbacks_follow_rows_and_java_listener_control_flow() -> Result<()> {
    let (_directory, path) = extra_workbook_fixture()?;
    let all_extras = HashSet::from([
        CellExtraType::Comment,
        CellExtraType::Hyperlink,
        CellExtraType::Merge,
    ]);
    let read_options = ReadOptions {
        sheet: SheetSelector::Name("Meta".to_owned()),
        extra_read: all_extras,
        custom_object: Some(CustomReadObject::new("reader-context".to_owned())),
        ..options()
    };
    let mut probe = ExtraProbe::default();
    read_xlsx::<TestRow, _>(&path, &read_options, &mut probe)?;
    assert_eq!(probe.extras.len(), 4);
    let first_extra = probe
        .events
        .iter()
        .position(|event| *event == "extra")
        .expect("extra event");
    assert!(
        probe.events[..first_extra]
            .iter()
            .all(|event| matches!(*event, "head" | "row"))
    );
    assert!(
        probe.events[first_extra..probe.events.len() - 1]
            .iter()
            .all(|event| *event == "extra")
    );
    assert_eq!(probe.events.last(), Some(&"after"));
    assert!(
        probe
            .context_customs
            .iter()
            .all(|value| value.as_deref() == Some("reader-context"))
    );

    let merge = probe
        .extras
        .iter()
        .find(|extra| extra.extra_type() == CellExtraType::Merge)
        .expect("merge extra");
    assert_eq!(merge.first_row_index(), 3);
    assert_eq!(merge.last_row_index(), 3);
    assert_eq!(merge.first_column_index(), 0);
    assert_eq!(merge.last_column_index(), 1);
    let hyperlinks = probe
        .extras
        .iter()
        .filter(|extra| extra.extra_type() == CellExtraType::Hyperlink)
        .filter_map(CellExtra::text)
        .collect::<Vec<_>>();
    assert!(hyperlinks.contains(&"https://example.com"));
    assert!(hyperlinks.contains(&"Meta!A1"));
    let comment = probe
        .extras
        .iter()
        .find(|extra| extra.extra_type() == CellExtraType::Comment)
        .expect("comment extra");
    assert_eq!(comment.text(), Some("Author:\ncomment & text"));
    assert_eq!(comment.first_row_index(), 1);
    assert_eq!(comment.first_column_index(), 0);

    let mut comments_only = ExtraProbe::default();
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            extra_read: HashSet::from([CellExtraType::Comment]),
            ..options()
        },
        &mut comments_only,
    )?;
    assert_eq!(comments_only.extras.len(), 1);
    assert_eq!(comments_only.extras[0].extra_type(), CellExtraType::Comment);

    let mut stopped = ExtraProbe {
        stop_after_extra: true,
        ..ExtraProbe::default()
    };
    read_xlsx::<TestRow, _>(&path, &read_options, &mut stopped)?;
    assert_eq!(stopped.extras.len(), 1);
    assert!(!stopped.events.contains(&"after"));

    let mut continued_error = ExtraProbe {
        fail_extra: true,
        error_action: Some(ErrorAction::Continue),
        ..ExtraProbe::default()
    };
    read_xlsx::<TestRow, _>(&path, &read_options, &mut continued_error)?;
    assert_eq!(continued_error.errors, 4);
    assert_eq!(continued_error.events.last(), Some(&"after"));
    assert!(
        continued_error
            .context_customs
            .iter()
            .all(|value| value.as_deref() == Some("reader-context"))
    );

    let mut stopped_error = ExtraProbe {
        fail_extra: true,
        ..ExtraProbe::default()
    };
    assert!(read_xlsx::<TestRow, _>(&path, &read_options, &mut stopped_error).is_err());
    assert_eq!(stopped_error.errors, 1);
    assert!(!stopped_error.events.contains(&"after"));

    let malformed = path.with_file_name("malformed-extra.xlsx");
    rewrite_first_sheet(&path, &malformed, "<worksheet>")?;
    let mut malformed_probe = ExtraProbe::default();
    assert!(read_xlsx::<TestRow, _>(&malformed, &read_options, &mut malformed_probe).is_err());
    Ok(())
}

#[test]
fn non_xlsx_readers_reject_requested_extra_metadata_before_opening_input() {
    let options = ReadOptions {
        extra_read: HashSet::from([CellExtraType::Comment]),
        ..options()
    };
    let mut probe = Probe::default();
    assert!(matches!(
        read_xls::<TestRow, _>(Path::new("missing.xls"), &options, &mut probe),
        Err(ExcelError::Unsupported(message)) if message.contains("XLS")
    ));
    assert!(matches!(
        read_csv::<TestRow, _>(Path::new("missing.csv"), &options, &mut probe),
        Err(ExcelError::Unsupported(message)) if message.contains("CSV")
    ));
}

#[test]
#[allow(clippy::too_many_lines)]
fn xlsx_stream_matches_java_cell_types_cached_formulas_dates_and_trimming() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("mixed-cells.xlsx");
    write_xlsx_package(
        &path,
        &[
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <workbookPr date1904="1"/>
  <sheets><sheet name="Mixed" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
            ),
            (
                "xl/sharedStrings.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><r><t xml:space="preserve"><![CDATA[  shared_x000D_]]></t></r><rPh><t>ignored</t></rPh><r><t xml:space="preserve">value  </t></r></si>
</sst>"#,
            ),
            (
                "xl/styles.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs>
</styleSheet>"#,
            ),
            (
                "xl/worksheets/sheet1.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:I2"/>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Shared</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Inline</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Boolean</t></is></c>
      <c r="D1" t="inlineStr"><is><t>Integer</t></is></c>
      <c r="E1" t="inlineStr"><is><t>Float</t></is></c>
      <c r="F1" t="inlineStr"><is><t>Formula number</t></is></c>
      <c r="G1" t="inlineStr"><is><t>Formula string</t></is></c>
      <c r="H1" t="inlineStr"><is><t>Error</t></is></c>
      <c r="I1" t="inlineStr"><is><t>Date</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>0</v></c>
      <c r="B2" t="inlineStr"><is><r><t xml:space="preserve"><![CDATA[  inline ]]></t></r><rPh><t>ignored</t></rPh><r><t xml:space="preserve">value  </t></r></is></c>
      <c r="C2" t="b"><v>1</v></c>
      <c r="D2" t="n"><v>42</v></c>
      <c r="E2" t="n"><v>3.5</v></c>
      <c r="F2"><f><![CDATA[SUM(D2:E2)]]></f><v>45.5</v></c>
      <c r="G2" t="str"><f>CONCAT("cache","d")</f><v>cached</v></c>
      <c r="H2" t="e"><v>#DIV/0!</v></c>
      <c r="I2" s="1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#,
            ),
        ],
    )?;

    let mut probe = RawProbe::default();
    read_xlsx::<RawRow, _>(
        &path,
        &ReadOptions {
            use_1904_windowing: true,
            ..options()
        },
        &mut probe,
    )?;
    assert_eq!(probe.0.len(), 1);
    assert_eq!(
        probe.0[0].cells[0],
        CellValue::String("shared\rvalue".to_owned())
    );
    assert_eq!(
        probe.0[0].cells[1],
        CellValue::String("inline value".to_owned())
    );
    assert_eq!(probe.0[0].cells[2], CellValue::Bool(true));
    assert_eq!(probe.0[0].cells[3], CellValue::Float(42.0));
    assert_eq!(probe.0[0].cells[4], CellValue::Float(3.5));
    assert_eq!(probe.0[0].cells[5], CellValue::Float(45.5));
    assert_eq!(probe.0[0].cells[6], CellValue::String("cached".to_owned()));
    assert_eq!(probe.0[0].cells[7], CellValue::String("#DIV/0!".to_owned()));
    assert_eq!(probe.0[0].cells[8].as_text(), "1904-01-02 00:00:00");
    assert_eq!(
        probe.0[0].formulas,
        vec![
            None,
            None,
            None,
            None,
            None,
            Some("SUM(D2:E2)".to_owned()),
            Some("CONCAT(\"cache\",\"d\")".to_owned()),
            None,
            None,
        ]
    );

    let expected = probe.0[0].clone();
    for read_cache in [ReadCacheMode::Memory, ReadCacheMode::Disk] {
        let mut cached = RawProbe::default();
        read_xlsx::<RawRow, _>(
            &path,
            &ReadOptions {
                use_1904_windowing: true,
                read_cache,
                ..options()
            },
            &mut cached,
        )?;
        assert_eq!(cached.0.as_slice(), std::slice::from_ref(&expected));
    }

    let mut untrimmed = RawProbe::default();
    read_xlsx::<RawRow, _>(
        &path,
        &ReadOptions {
            auto_trim: false,
            ..options()
        },
        &mut untrimmed,
    )?;
    assert_eq!(
        untrimmed.0[0].cells[0],
        CellValue::String("  shared\rvalue  ".to_owned())
    );
    assert_eq!(
        untrimmed.0[0].cells[1],
        CellValue::String("  inline value  ".to_owned())
    );

    let mut java_default = DynamicProbe::default();
    read_xlsx::<DynamicRow, _>(&path, &options(), &mut java_default)?;
    assert_eq!(
        java_default.0[0].get(8),
        Some(&DynamicValue::String("1/1/00".to_owned()))
    );

    let mut strings = DynamicProbe::default();
    read_xlsx::<DynamicRow, _>(
        &path,
        &ReadOptions {
            use_1904_windowing: true,
            ..options()
        },
        &mut strings,
    )?;
    assert_eq!(
        strings.0[0].get(3),
        Some(&DynamicValue::String("42".to_owned()))
    );
    assert_eq!(
        strings.0[0].get(8),
        Some(&DynamicValue::String("1/2/04".to_owned()))
    );

    let mut actual = DynamicProbe::default();
    read_xlsx::<DynamicRow, _>(
        &path,
        &ReadOptions {
            read_default_return: ReadDefaultReturn::ActualData,
            use_1904_windowing: true,
            ..options()
        },
        &mut actual,
    )?;
    assert_eq!(
        actual.0[0].get(2),
        Some(&DynamicValue::ActualData(CellValue::Bool(true)))
    );
    assert_eq!(
        actual.0[0].get(5),
        Some(&DynamicValue::ActualData(CellValue::Decimal(
            "45.5".parse().map_err(test_error)?
        )))
    );
    assert_eq!(
        actual.0[0].get(7),
        Some(&DynamicValue::ActualData(CellValue::String(
            "#DIV/0!".to_owned()
        )))
    );

    let mut cell_data = DynamicProbe::default();
    read_xlsx::<DynamicRow, _>(
        &path,
        &ReadOptions {
            read_default_return: ReadDefaultReturn::ReadCellData,
            use_1904_windowing: true,
            ..options()
        },
        &mut cell_data,
    )?;
    let DynamicValue::ReadCellData(formula_cell) =
        cell_data.0[0].get(5).expect("formula cell data")
    else {
        panic!("expected formula read cell data");
    };
    let expected_decimal = CellValue::Decimal("45.5".parse().map_err(test_error)?);
    assert_eq!(formula_cell.raw_value(), &expected_decimal);
    assert_eq!(formula_cell.data(), &expected_decimal);
    assert_eq!(
        formula_cell.formula().map(FormulaData::formula_value),
        Some("SUM(D2:E2)")
    );
    Ok(())
}

#[test]
fn xlsx_primary_cell_stream_rejects_malformed_xml() -> Result<()> {
    let (directory, base) = workbook_fixture()?;
    let cases = [
        (
            "display-cell-xml-error.xlsx",
            r#"<worksheet><sheetData><row r="1"><c r="A1"><v><"#,
        ),
        (
            "display-tail-xml-error.xlsx",
            r#"<worksheet><sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>one</t></is></c></row><"#,
        ),
    ];

    for (name, replacement) in cases {
        let metadata_path = directory.path().join(name);
        rewrite_first_sheet(&base, &metadata_path, replacement)?;
        let mut listener = DynamicProbe::default();
        assert!(read_xlsx::<DynamicRow, _>(&metadata_path, &options(), &mut listener).is_err());
    }
    Ok(())
}

#[test]
fn dynamic_xlsx_reports_display_stream_initialization_errors() -> Result<()> {
    let (directory, base) = workbook_fixture()?;
    let malformed = directory.path().join("missing-sheet-data.xlsx");
    rewrite_first_sheet(&base, &malformed, "<worksheet/>")?;
    let mut listener = DynamicProbe::default();
    assert!(read_xlsx::<DynamicRow, _>(&malformed, &options(), &mut listener).is_err());
    Ok(())
}

#[test]
fn dynamic_rows_preserve_xlsx_gaps_and_csv_scalar_contracts() -> Result<()> {
    let (directory, base) = workbook_fixture()?;
    let sparse = directory.path().join("dynamic-sparse.xlsx");
    rewrite_first_sheet(
        &base,
        &sparse,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>First</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Tail</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>value</t></is></c>
      <c r="C2" t="n"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>"#,
    )?;
    let mut xlsx = DynamicProbe::default();
    read_xlsx::<DynamicRow, _>(&sparse, &options(), &mut xlsx)?;
    assert_eq!(xlsx.0[0].values().len(), 3);
    assert_eq!(
        xlsx.0[0].get(0),
        Some(&DynamicValue::String("value".to_owned()))
    );
    assert_eq!(xlsx.0[0].get(1), Some(&DynamicValue::Null));
    assert_eq!(
        xlsx.0[0].get(2),
        Some(&DynamicValue::String("3".to_owned()))
    );

    let csv_path = directory.path().join("dynamic.csv");
    fs::write(&csv_path, "Text,Number,Empty,Tail\r\nvalue,109,,last\r\n")?;
    let mut csv_strings = DynamicProbe::default();
    read_csv::<DynamicRow, _>(&csv_path, &options(), &mut csv_strings)?;
    assert_eq!(
        csv_strings.0[0].get(1),
        Some(&DynamicValue::String("109".to_owned()))
    );
    assert_eq!(
        csv_strings.0[0].get(2),
        Some(&DynamicValue::String(String::new()))
    );

    let mut csv_actual = DynamicProbe::default();
    read_csv::<DynamicRow, _>(
        &csv_path,
        &ReadOptions {
            read_default_return: ReadDefaultReturn::ActualData,
            ..options()
        },
        &mut csv_actual,
    )?;
    assert_eq!(
        csv_actual.0[0].get(1),
        Some(&DynamicValue::ActualData(CellValue::String(
            "109".to_owned()
        )))
    );
    assert_eq!(
        csv_actual.0[0].get(2),
        Some(&DynamicValue::ActualData(CellValue::String(String::new())))
    );
    Ok(())
}

#[test]
fn sheet_selection_supports_first_index_name_all_and_missing_values() -> Result<()> {
    let (_directory, path) = workbook_fixture()?;
    let workbook: Xlsx<_> = open_workbook(path).map_err(test_error)?;
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::First, true)?,
        vec![(0, "First".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::Index(1), true)?,
        vec![(1, "Second".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::Name("Second".to_owned()), true,)?,
        vec![(1, "Second".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::All, true)?.len(),
        2
    );
    assert!(selected_sheet_names(&workbook, &SheetSelector::Index(2), true).is_err());
    assert!(
        selected_sheet_names(&workbook, &SheetSelector::Name("Missing".to_owned()), true,).is_err()
    );
    assert!(select_sheet_names(Vec::new(), &SheetSelector::First, true).is_err());
    assert_eq!(
        select_sheet_names(
            vec![" First ".to_owned()],
            &SheetSelector::Name("First".to_owned()),
            true,
        )?,
        vec![(0, " First ".to_owned())]
    );
    assert!(
        select_sheet_names(
            vec![" First ".to_owned()],
            &SheetSelector::Name("First".to_owned()),
            false,
        )
        .is_err()
    );

    let legacy = || {
        vec![
            ("First".to_owned(), Range::empty()),
            ("Second".to_owned(), Range::empty()),
        ]
    };
    let first = select_xls_sheets(legacy(), &SheetSelector::First, true)?;
    assert_eq!((first[0].0, first[0].1.as_str()), (0, "First"));
    let second = select_xls_sheets(legacy(), &SheetSelector::Index(1), true)?;
    assert_eq!((second[0].0, second[0].1.as_str()), (1, "Second"));
    let named = select_xls_sheets(legacy(), &SheetSelector::Name("Second".to_owned()), true)?;
    assert_eq!((named[0].0, named[0].1.as_str()), (1, "Second"));
    assert_eq!(
        select_xls_sheets(legacy(), &SheetSelector::All, true)?.len(),
        2
    );
    assert!(select_xls_sheets(legacy(), &SheetSelector::Index(2), true).is_err());
    assert!(
        select_xls_sheets(legacy(), &SheetSelector::Name("Missing".to_owned()), true,).is_err()
    );
    assert!(select_xls_sheets(Vec::new(), &SheetSelector::First, true).is_err());
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn csv_read_uses_typed_lifecycle_single_sheet_selection_and_flexible_rows() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("fixture.csv");
    fs::write(&path, b"\xEF\xBB\xBFValue,Extra\r\none,1\r\ntwo\r\n")?;
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_csv::<TestRow, _>(&path, &options(), &mut probe)?;
    assert_eq!(
        probe.rows,
        vec![TestRow("one".to_owned()), TestRow("two".to_owned())]
    );
    assert_eq!(probe.heads[0].get("Value"), Some(&0));
    assert_eq!(probe.after, vec![("Sheet1".to_owned(), 0, 2)]);

    assert_eq!(csv_sheet_name(&SheetSelector::First)?, "Sheet1");
    assert_eq!(csv_sheet_name(&SheetSelector::Index(0))?, "Sheet1");
    assert_eq!(csv_sheet_name(&SheetSelector::All)?, "Sheet1");
    assert_eq!(
        csv_sheet_name(&SheetSelector::Name("Custom".to_owned()))?,
        "Custom"
    );
    assert!(csv_sheet_name(&SheetSelector::Index(1)).is_err());
    assert_eq!(csv_row_index(0)?, 0);
    if usize::BITS > 32 {
        assert!(csv_row_index(usize::try_from(u64::from(u32::MAX) + 1).unwrap()).is_err());
    }

    let malformed_utf8 = directory.path().join("malformed-utf8.csv");
    fs::write(&malformed_utf8, [0xff])?;
    let mut replacement_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_csv::<TestRow, _>(
        &malformed_utf8,
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut replacement_probe,
    )?;
    assert_eq!(replacement_probe.rows, vec![TestRow("�".to_owned())]);
    assert!(
        read_csv::<TestRow, _>(
            &path,
            &ReadOptions {
                sheet: SheetSelector::Index(1),
                ..options()
            },
            &mut probe
        )
        .is_err()
    );
    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(read_csv::<TestRow, _>(&path, &options(), &mut failing_head).is_err());
    let record = csv::StringRecord::from(vec!["value"]);
    let mut record_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_csv_records::<TestRow, _>(
        &mut [Ok(record.clone())].into_iter(),
        0,
        "Sheet1",
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut record_probe,
    )?;
    assert_eq!(record_probe.rows, vec![TestRow("value".to_owned())]);
    read_csv_records::<TestRow, _>(
        &mut [Ok(record.clone()), Ok(record.clone())].into_iter(),
        0,
        "Sheet1",
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut record_probe,
    )?;
    assert_eq!(record_probe.rows.len(), 3);
    let mut stopped = Probe::default();
    read_csv_records::<TestRow, _>(
        &mut [Ok(record.clone()), Ok(record.clone())].into_iter(),
        0,
        "Sheet1",
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut stopped,
    )?;
    assert_eq!(stopped.rows, vec![TestRow("value".to_owned())]);
    assert!(stopped.after.is_empty());
    let mut invalid_utf8_reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_reader([0xFF].as_slice());
    assert!(
        read_csv_records::<TestRow, _>(
            &mut invalid_utf8_reader.records(),
            0,
            "Sheet1",
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut record_probe,
        )
        .is_err()
    );
    if usize::BITS > 32 {
        assert!(
            read_csv_records::<TestRow, _>(
                &mut [Ok(record.clone())].into_iter(),
                usize::try_from(u64::from(u32::MAX) + 1).unwrap(),
                "Sheet1",
                &ReadOptions {
                    head_row_number: 0,
                    ..options()
                },
                &mut probe
            )
            .is_err()
        );
    }
    assert!(
        read_csv_records::<TestRow, _>(
            &mut [Ok(record.clone()), Ok(record)].into_iter(),
            usize::MAX,
            "Sheet1",
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut probe
        )
        .is_err()
    );
    assert!(
        read_csv::<TestRow, _>(
            &directory.path().join("missing.csv"),
            &options(),
            &mut probe
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn csv_read_decodes_java_charset_names_and_strips_matching_bom() -> Result<()> {
    let directory = tempdir()?;
    for (name, encoding, bom) in [
        ("utf-8", encoding_rs::UTF_8, b"\xEF\xBB\xBF".as_slice()),
        ("GBK", encoding_rs::GBK, b"".as_slice()),
        ("UTF-16BE", encoding_rs::UTF_16BE, b"\xFE\xFF".as_slice()),
        ("UTF-16LE", encoding_rs::UTF_16LE, b"\xFF\xFE".as_slice()),
    ] {
        let path = directory
            .path()
            .join(format!("{}.csv", name.to_lowercase()));
        let mut bytes = bom.to_vec();
        bytes.extend_from_slice(&encode_csv_fixture(encoding, "Value\r\n姓名\r\n"));
        fs::write(&path, bytes)?;

        let mut probe = Probe {
            continue_reading: true,
            ..Probe::default()
        };
        read_csv::<TestRow, _>(
            &path,
            &ReadOptions {
                charset: CsvCharset::new(name),
                ..options()
            },
            &mut probe,
        )?;
        assert_eq!(
            probe.rows,
            vec![TestRow("姓名".to_owned())],
            "charset {name}"
        );
    }

    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    let error = read_csv::<TestRow, _>(
        &directory.path().join("utf-8.csv"),
        &ReadOptions {
            charset: CsvCharset::new("not-a-charset"),
            ..options()
        },
        &mut probe,
    )
    .expect_err("unknown charset must be rejected");
    assert!(matches!(error, ExcelError::Unsupported(_)));
    Ok(())
}

#[test]
fn reads_java_easyexcel_official_csv_bom_fixtures() -> Result<()> {
    let directory = tempdir()?;
    for (name, fixture) in [
        (
            "no-bom.csv",
            include_str!("fixtures/java-bom-no-bom.csv.b64"),
        ),
        (
            "office-bom.csv",
            include_str!("fixtures/java-bom-office-bom.csv.b64"),
        ),
    ] {
        let path = directory.path().join(name);
        fs::write(
            &path,
            base64::engine::general_purpose::STANDARD
                .decode(fixture.trim())
                .map_err(test_error)?,
        )?;
        let mut probe = Probe {
            continue_reading: true,
            ..Probe::default()
        };
        read_csv::<TestRow, _>(&path, &options(), &mut probe)?;
        assert_eq!(probe.rows.len(), 10);
        assert_eq!(probe.rows[0], TestRow("姓名0".to_owned()));
        assert_eq!(probe.rows[9], TestRow("姓名9".to_owned()));
    }
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn row_processing_handles_headers_skips_data_and_listener_failures() -> Result<()> {
    let mut headers = Arc::new(HashMap::new());
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    process_row::<TestRow>(
        0,
        "First",
        0,
        vec![CellValue::String(" Value ".to_owned()), CellValue::Empty],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.heads[0].get("Value"), Some(&0));

    process_row::<TestRow>(
        0,
        "First",
        1,
        vec![CellValue::String("one".to_owned())],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows, vec![TestRow("one".to_owned())]);

    process_row::<TestRow>(
        0,
        "First",
        2,
        vec![CellValue::Empty],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows.len(), 1);

    let two_header_rows = ReadOptions {
        head_row_number: 2,
        ..options()
    };
    assert_eq!(
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("ignored".to_owned())],
            &two_header_rows,
            &mut headers,
            &mut probe,
        )?,
        ReadFlow::Continue
    );
    assert_eq!(probe.heads.len(), 2);
    process_row::<TestRow>(
        0,
        "First",
        1,
        vec![CellValue::String("Final".to_owned())],
        &two_header_rows,
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.heads.len(), 3);
    assert_eq!(headers.get("Final"), Some(&0));
    assert_eq!(probe.rows.len(), 1);

    probe.continue_reading = false;
    let include_empty = ReadOptions {
        ignore_empty_row: false,
        ..options()
    };
    assert_eq!(
        process_row::<TestRow>(
            0,
            "First",
            3,
            vec![CellValue::Empty],
            &include_empty,
            &mut headers,
            &mut probe,
        )?,
        ReadFlow::Stop
    );
    assert_eq!(probe.rows.len(), 2);

    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("Value".to_owned())],
            &options(),
            &mut headers,
            &mut failing_head
        )
        .is_err()
    );
    assert_eq!(failing_head.errors, 1);

    let no_head = ReadOptions {
        head_row_number: 0,
        ..options()
    };
    let mut tolerated_invoke = Probe {
        continue_reading: true,
        fail_invoke: true,
        error_action: Some(ErrorAction::Continue),
        ..Probe::default()
    };
    assert_eq!(
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("value".to_owned())],
            &no_head,
            &mut headers,
            &mut tolerated_invoke,
        )?,
        ReadFlow::Continue
    );
    assert_eq!(tolerated_invoke.errors, 1);

    let mut trimming_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    process_row::<TestRow>(
        0,
        "First",
        0,
        vec![CellValue::String("  trimmed  ".to_owned())],
        &no_head,
        &mut headers,
        &mut trimming_probe,
    )?;
    assert_eq!(trimming_probe.rows, vec![TestRow("trimmed".to_owned())]);
    process_row::<TestRow>(
        0,
        "First",
        1,
        vec![CellValue::String("   ".to_owned())],
        &no_head,
        &mut headers,
        &mut trimming_probe,
    )?;
    assert_eq!(trimming_probe.rows.len(), 1);

    let mut untrimmed_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    process_row::<TestRow>(
        0,
        "First",
        0,
        vec![CellValue::String("  preserved  ".to_owned())],
        &ReadOptions {
            head_row_number: 0,
            auto_trim: false,
            ..options()
        },
        &mut headers,
        &mut untrimmed_probe,
    )?;
    assert_eq!(
        untrimmed_probe.rows,
        vec![TestRow("  preserved  ".to_owned())]
    );
    Ok(())
}

#[test]
fn conversion_error_actions_continue_skip_or_stop() -> Result<()> {
    let mut headers = Arc::new(HashMap::new());
    let read_options = ReadOptions {
        head_row_number: 0,
        ignore_empty_row: false,
        ..options()
    };
    for action in [ErrorAction::Continue, ErrorAction::SkipRow] {
        let mut listener = ErrorProbe { action, errors: 0 };
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("conversion-error".to_owned())],
            &read_options,
            &mut headers,
            &mut listener,
        )?;
        assert_eq!(listener.errors, 1);
    }
    let mut listener = ErrorProbe {
        action: ErrorAction::Stop,
        errors: 0,
    };
    assert!(
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("conversion-error".to_owned())],
            &read_options,
            &mut headers,
            &mut listener
        )
        .is_err()
    );
    assert_eq!(listener.errors, 1);
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn public_reader_streams_all_sheets_and_reports_invalid_workbooks() -> Result<()> {
    let (fixture_directory, path) = workbook_fixture()?;
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            sheet: SheetSelector::All,
            ..options()
        },
        &mut probe,
    )?;
    assert_eq!(
        probe.rows,
        vec![TestRow("one".to_owned()), TestRow("two".to_owned())]
    );
    assert_eq!(
        probe.after,
        vec![("First".to_owned(), 0, 1), ("Second".to_owned(), 1, 1)]
    );

    let mut failing_after = Probe {
        continue_reading: true,
        fail_after: true,
        ..Probe::default()
    };
    assert!(read_xlsx::<TestRow, _>(&path, &options(), &mut failing_after).is_err());

    let mut stopped = Probe::default();
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            sheet: SheetSelector::All,
            ..options()
        },
        &mut stopped,
    )?;
    assert_eq!(stopped.heads.len(), 1);
    assert!(stopped.rows.is_empty());
    assert!(stopped.after.is_empty());

    let mut missing = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &path,
            &ReadOptions {
                sheet: SheetSelector::Index(99),
                ..options()
            },
            &mut missing,
        )
        .is_err()
    );

    let mut failing_transition = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(&path, &options(), &mut failing_transition).is_err(),
        "a header error emitted while advancing rows must propagate"
    );

    let single_path = fixture_directory.path().join("single.xlsx");
    let mut workbook = Workbook::new();
    workbook
        .add_worksheet()
        .write_string(0, 0, "Value")
        .map_err(test_error)?;
    workbook.save(&single_path).map_err(test_error)?;
    let mut failing_final = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(&single_path, &options(), &mut failing_final).is_err(),
        "a header error emitted at end-of-sheet must propagate"
    );
    let mut stopped_final = Probe::default();
    read_xlsx::<TestRow, _>(
        &single_path,
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut stopped_final,
    )?;
    assert_eq!(stopped_final.rows, vec![TestRow("Value".to_owned())]);
    assert!(stopped_final.after.is_empty());

    let source = XlsxSource::open(&path, None)?;
    let mut metadata = XlsxRowMetadata::new(source.reader()?)?;
    assert!(
        metadata
            .display_cells("Missing", false, false, ssfmt::Locale::default())
            .is_err()
    );

    let empty_path = fixture_directory.path().join("empty.xlsx");
    let mut empty_workbook = Workbook::new();
    empty_workbook.add_worksheet();
    empty_workbook.save(&empty_path).map_err(test_error)?;
    let mut empty_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(&empty_path, &options(), &mut empty_probe)?;
    assert!(empty_probe.rows.is_empty());
    assert_eq!(empty_probe.after, vec![("Sheet1".to_owned(), 0, 0)]);

    let out_of_order_path = fixture_directory.path().join("out-of-order.xlsx");
    let out_of_order_xml = worksheet_xml(
        r#"<c r="B1" t="inlineStr"><is><t>second</t></is></c>
<c r="A1" t="inlineStr"><is><t>first</t></is></c>"#,
    );
    rewrite_first_sheet(&path, &out_of_order_path, &out_of_order_xml)?;
    let mut out_of_order_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &out_of_order_path,
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut out_of_order_probe,
    )?;
    assert_eq!(out_of_order_probe.rows, vec![TestRow("first".to_owned())]);

    let sparse_path = fixture_directory.path().join("sparse.xlsx");
    let sparse_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>one</t></is></c></row>
  </sheetData>
</worksheet>"#;
    rewrite_first_sheet(&path, &sparse_path, sparse_xml)?;
    let sparse_options = ReadOptions {
        ignore_empty_row: false,
        ..options()
    };
    let mut sparse_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(&sparse_path, &sparse_options, &mut sparse_probe)?;
    assert_eq!(
        sparse_probe.rows,
        vec![
            TestRow(String::new()),
            TestRow(String::new()),
            TestRow("one".to_owned())
        ]
    );

    let mut stopped_sparse = Probe {
        continue_reading: true,
        stop_after_callbacks: Some(2),
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(&sparse_path, &sparse_options, &mut stopped_sparse)?;
    assert_eq!(stopped_sparse.rows, vec![TestRow(String::new())]);
    assert!(stopped_sparse.after.is_empty());

    let mut failing_sparse = Probe {
        continue_reading: true,
        fail_invoke: true,
        ..Probe::default()
    };
    assert!(read_xlsx::<TestRow, _>(&sparse_path, &sparse_options, &mut failing_sparse).is_err());
    assert_eq!(failing_sparse.errors, 1);

    let trailing_empty_path = fixture_directory.path().join("trailing-empty.xlsx");
    let trailing_empty_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>one</t></is></c></row>
    <row r="5"/>
  </sheetData>
</worksheet>"#;
    rewrite_first_sheet(&path, &trailing_empty_path, trailing_empty_xml)?;
    let mut trailing_empty_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &trailing_empty_path,
        &sparse_options,
        &mut trailing_empty_probe,
    )?;
    assert_eq!(
        trailing_empty_probe.rows,
        vec![
            TestRow("one".to_owned()),
            TestRow(String::new()),
            TestRow(String::new()),
            TestRow(String::new())
        ]
    );
    assert_eq!(trailing_empty_probe.after, vec![("First".to_owned(), 0, 4)]);

    let mut stopped_trailing = Probe {
        continue_reading: true,
        stop_after_callbacks: Some(3),
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(&trailing_empty_path, &sparse_options, &mut stopped_trailing)?;
    assert_eq!(
        stopped_trailing.rows,
        vec![TestRow("one".to_owned()), TestRow(String::new())]
    );
    assert!(stopped_trailing.after.is_empty());

    let mut failing_trailing = Probe {
        continue_reading: true,
        fail_invoke_at: Some(2),
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(&trailing_empty_path, &sparse_options, &mut failing_trailing,)
            .is_err()
    );
    assert_eq!(failing_trailing.errors, 1);

    let empty_rows_path = fixture_directory.path().join("only-empty-rows.xlsx");
    let empty_rows_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row/><row/><row/></sheetData>
</worksheet>"#;
    rewrite_first_sheet(&path, &empty_rows_path, empty_rows_xml)?;
    let mut empty_rows_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &empty_rows_path,
        &ReadOptions {
            head_row_number: 0,
            ignore_empty_row: false,
            ..options()
        },
        &mut empty_rows_probe,
    )?;
    assert_eq!(
        empty_rows_probe.rows,
        vec![
            TestRow(String::new()),
            TestRow(String::new()),
            TestRow(String::new())
        ]
    );
    assert_eq!(empty_rows_probe.after, vec![("First".to_owned(), 0, 2)]);

    let invalid_row_path = fixture_directory.path().join("invalid-row.xlsx");
    let invalid_row_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="0"/></sheetData>
</worksheet>"#;
    rewrite_first_sheet(&path, &invalid_row_path, invalid_row_xml)?;
    let mut invalid_row_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &invalid_row_path,
            &ReadOptions {
                ignore_empty_row: false,
                ..options()
            },
            &mut invalid_row_probe,
        )
        .is_err()
    );

    let missing_sheet_path = fixture_directory.path().join("missing-sheet-part.xlsx");
    remove_first_sheet(&path, &missing_sheet_path)?;
    let mut missing_sheet_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &missing_sheet_path,
            &ReadOptions {
                ignore_empty_row: false,
                ..options()
            },
            &mut missing_sheet_probe,
        )
        .is_err()
    );

    let leading_sparse_path = fixture_directory.path().join("leading-sparse.xlsx");
    let leading_sparse_xml = worksheet_xml(r#"<c r="A3" t="inlineStr"><is><t>first</t></is></c>"#)
        .replace("<row r=\"1\">", "<row r=\"3\">");
    rewrite_first_sheet(&path, &leading_sparse_path, &leading_sparse_xml)?;
    let mut leading_sparse_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &leading_sparse_path,
        &ReadOptions {
            head_row_number: 0,
            ignore_empty_row: false,
            ..options()
        },
        &mut leading_sparse_probe,
    )?;
    assert_eq!(
        leading_sparse_probe.rows,
        vec![
            TestRow(String::new()),
            TestRow(String::new()),
            TestRow("first".to_owned())
        ]
    );

    let wide_path = fixture_directory.path().join("wide.xlsx");
    let wide_column = column_name(u32::from(u16::MAX) + 1);
    let wide_xml = worksheet_xml(&format!(
        r#"<c r="{wide_column}1" t="inlineStr"><is><t>wide</t></is></c>"#
    ));
    rewrite_first_sheet(&path, &wide_path, &wide_xml)?;
    let mut wide_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &wide_path,
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut wide_probe,
        )
        .is_err()
    );

    let truncated_path = fixture_directory.path().join("truncated.xlsx");
    rewrite_first_sheet(
        &path,
        &truncated_path,
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c>"#,
    )?;
    let mut truncated_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &truncated_path,
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut truncated_probe,
        )
        .is_err()
    );

    let directory = tempdir()?;
    let invalid = directory.path().join("invalid.xlsx");
    fs::write(&invalid, b"not an xlsx")?;
    assert!(read_xlsx::<TestRow, _>(&invalid, &options(), &mut probe).is_err());
    assert!(
        read_xlsx::<TestRow, _>(
            &directory.path().join("missing.xlsx"),
            &options(),
            &mut probe,
        )
        .is_err()
    );
    let missing_source = XlsxSource::File(directory.path().join("missing-source.xlsx"));
    assert!(read_xlsx_source::<TestRow, _>(&missing_source, &options(), &mut probe).is_err());
    assert!(read_xlsx::<TestRow, _>(directory.path(), &options(), &mut probe).is_err());
    let invalid_encrypted = directory.path().join("invalid-encrypted.xlsx");
    fs::write(
        &invalid_encrypted,
        [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
    )?;
    assert!(
        read_xlsx::<TestRow, _>(
            &invalid_encrypted,
            &ReadOptions {
                password: Some("123456".to_owned()),
                ..options()
            },
            &mut probe,
        )
        .is_err()
    );
    Ok(())
}
