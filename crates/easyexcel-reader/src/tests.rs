use std::collections::HashMap;
use std::fs;
use std::io::{Read, Write};
use std::sync::Arc;

use calamine::{CellErrorType, ExcelDateTime, ExcelDateTimeType};
use easyexcel_core::{ExcelColumn, IntoExcelCell};
use rust_xlsxwriter::Workbook;
use tempfile::{TempDir, tempdir};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use super::*;

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

#[derive(Default)]
struct Probe {
    heads: Vec<HashMap<String, usize>>,
    rows: Vec<TestRow>,
    after: Vec<(String, usize, u32)>,
    continue_reading: bool,
    fail_head: bool,
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
        self.rows.push(data);
        Ok(())
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.after.push((
            context.sheet_name().to_owned(),
            context.sheet_no(),
            context.row_index(),
        ));
        Ok(())
    }

    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        self.continue_reading
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
            CellValue::Error("Div0".to_owned()),
        ),
    ];
    for (input, expected) in cases {
        assert_eq!(from_calamine(&input), expected);
    }
    assert!(matches!(
        from_calamine(&DataRef::DateTime(datetime)),
        CellValue::DateTime(_)
    ));
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
}

#[test]
fn sheet_selection_supports_first_index_name_all_and_missing_values() -> Result<()> {
    let (_directory, path) = workbook_fixture()?;
    let workbook: Xlsx<_> = open_workbook(path).map_err(test_error)?;
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::First)?,
        vec![(0, "First".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::Index(1))?,
        vec![(1, "Second".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::Name("Second".to_owned()))?,
        vec![(1, "Second".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::All)?.len(),
        2
    );
    assert!(selected_sheet_names(&workbook, &SheetSelector::Index(2)).is_err());
    assert!(selected_sheet_names(&workbook, &SheetSelector::Name("Missing".to_owned())).is_err());
    assert!(select_sheet_names(Vec::new(), &SheetSelector::First).is_err());
    Ok(())
}

#[test]
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
        vec![CellValue::String("Value".to_owned()), CellValue::Empty],
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
    process_row::<TestRow>(
        0,
        "First",
        0,
        vec![CellValue::String("ignored".to_owned())],
        &two_header_rows,
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows.len(), 1);

    probe.continue_reading = false;
    let include_empty = ReadOptions {
        ignore_empty_row: false,
        ..options()
    };
    process_row::<TestRow>(
        0,
        "First",
        3,
        vec![CellValue::Empty],
        &include_empty,
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows.len(), 1);

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

    let mut opened: Xlsx<_> = open_workbook(&path).map_err(test_error)?;
    let mut direct = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    let mut consumer = TypedRowConsumer::<TestRow> {
        listener: &mut direct,
    };
    assert!(read_sheet(&mut opened, 0, "Missing", &options(), &mut consumer).is_err());

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
    Ok(())
}
