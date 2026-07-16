use std::collections::HashMap;
use std::fs;
use std::sync::Arc;

use calamine::{CellErrorType, ExcelDateTime, ExcelDateTimeType};
use easyexcel_core::{ExcelColumn, IntoExcelCell};
use rust_xlsxwriter::Workbook;
use tempfile::{TempDir, tempdir};

use super::*;

#[derive(Debug, PartialEq, Eq)]
struct TestRow(String);

impl ExcelRow for TestRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Value", Some(0), 0, None)];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        Ok(Self(
            row.cell(&Self::schema()[0])
                .map_or_else(String::new, CellValue::as_text),
        ))
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

struct FailingRow;

impl ExcelRow for FailingRow {
    fn schema() -> &'static [ExcelColumn] {
        &[]
    }

    fn from_row(_row: &RowData) -> Result<Self> {
        Err(ExcelError::Format("conversion failed".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(Vec::new())
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

impl ReadListener<FailingRow> for ErrorProbe {
    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        self.errors += 1;
        self.action
    }

    fn invoke(&mut self, _data: FailingRow, _context: &AnalysisContext) -> Result<()> {
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
    first.set_name("First").map_err(format_error)?;
    first.write_string(0, 0, "Value").map_err(format_error)?;
    first.write_string(1, 0, "one").map_err(format_error)?;
    let second = workbook.add_worksheet();
    second.set_name("Second").map_err(format_error)?;
    second.write_string(0, 0, "Value").map_err(format_error)?;
    second.write_string(1, 0, "two").map_err(format_error)?;
    workbook.save(&path).map_err(format_error)?;
    Ok((directory, path))
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
    let workbook: Xlsx<_> = open_workbook(path).map_err(format_error)?;
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
    process_row::<TestRow, _>(
        0,
        "First",
        0,
        vec![CellValue::String("Value".to_owned()), CellValue::Empty],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.heads[0].get("Value"), Some(&0));

    process_row::<TestRow, _>(
        0,
        "First",
        1,
        vec![CellValue::String("one".to_owned())],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows, vec![TestRow("one".to_owned())]);

    process_row::<TestRow, _>(
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
    process_row::<TestRow, _>(
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
    process_row::<TestRow, _>(
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
        process_row::<TestRow, _>(
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
        process_row::<FailingRow, _>(
            0,
            "First",
            0,
            vec![CellValue::String("bad".to_owned())],
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
        process_row::<FailingRow, _>(
            0,
            "First",
            0,
            vec![CellValue::String("bad".to_owned())],
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
        .map_err(format_error)?;
    workbook.save(&single_path).map_err(format_error)?;
    let mut failing_final = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(&single_path, &options(), &mut failing_final).is_err(),
        "a header error emitted at end-of-sheet must propagate"
    );

    let mut opened: Xlsx<_> = open_workbook(&path).map_err(format_error)?;
    let mut direct = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_sheet::<TestRow, _, _>(&mut opened, 0, "Missing", &options(), &mut direct).is_err()
    );

    let directory = tempdir()?;
    let invalid = directory.path().join("invalid.xlsx");
    fs::write(&invalid, b"not an xlsx")?;
    assert!(read_xlsx::<TestRow, _>(&invalid, &options(), &mut probe).is_err());
    Ok(())
}
