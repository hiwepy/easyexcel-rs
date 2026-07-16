use std::cell::Cell;
use std::cell::RefCell;
use std::rc::Rc;

use calamine::{Data, Reader, Xlsx, open_workbook};
use chrono::NaiveDate;
use tempfile::tempdir;

use super::*;

fn test_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[derive(Clone)]
struct EveryCell {
    cells: Vec<CellValue>,
    fail: bool,
}

thread_local! {
    static USE_WIDE_SCHEMA: Cell<bool> = const { Cell::new(false) };
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
        ];
        const WIDE_COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("wide", "Wide", Some(65_536), 0, None)];
        USE_WIDE_SCHEMA.with(|wide| if wide.get() { WIDE_COLUMNS } else { COLUMNS })
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
        ],
        fail: false,
    }
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
            constant_memory: false,
            need_head: true,
            freeze_head: false,
            freeze_panes: None,
            include_column_indexes: None,
            include_column_field_names: None,
            exclude_column_indexes: Vec::new(),
            exclude_column_field_names: Vec::new(),
            order_by_include_column: false,
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
}

#[test]
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
    Ok(())
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
fn stateful_writer_supports_multiple_sheets_and_idempotent_finish() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("multi.xlsx");
    let events = Rc::new(RefCell::new(Vec::new()));
    let handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(RecordingHandler {
        order: 5,
        events: Rc::clone(&events),
    })];
    let first = WriteSheet::<EveryCell>::new("Users").freeze_head(true);
    let second = WriteSheet::<EveryCell>::new("Archive")
        .need_head(false)
        .constant_memory(true);
    assert_eq!(first.options().sheet_name, "Users");
    assert!(first.options().freeze_head);
    assert!(!second.options().need_head);
    assert!(second.options().constant_memory);

    let mut writer = ExcelWriter::with_handlers(&path, handlers);
    assert!(!writer.is_finished());
    writer
        .write(vec![every_cell()], &first)?
        .write(vec![every_cell(), every_cell()], &second)?;
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
    let users = workbook.worksheet_range("Users").map_err(test_error)?;
    assert_eq!(
        users.get_value((1, 1)),
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

    let mut duplicate = ExcelWriter::new(directory.path().join("duplicate.xlsx"));
    duplicate.write(Vec::new(), &sheet)?;
    assert!(duplicate.write(Vec::new(), &sheet).is_err());

    let invalid = WriteSheet::<EveryCell>::new("bad/name");
    let mut invalid_sheet = ExcelWriter::new(directory.path().join("invalid.xlsx"));
    assert!(invalid_sheet.write(Vec::new(), &invalid).is_err());

    let mut invalid_output = ExcelWriter::new(directory.path());
    assert!(invalid_output.finish().is_err());
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
            &mut handlers,
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
        write_xlsx::<EveryCell, _>(directory.path(), &WriteOptions::default(), Vec::new()).is_err()
    );
    Ok(())
}
