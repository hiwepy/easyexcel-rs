use calamine::{Data, Reader, Xlsx, open_workbook};
use chrono::NaiveDate;
use tempfile::tempdir;

use super::*;

#[derive(Clone)]
struct EveryCell {
    cells: Vec<CellValue>,
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
        COLUMNS
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Err(ExcelError::Unsupported("test-only writer row".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(self.cells.clone())
    }
}

struct BrokenRow;

impl ExcelRow for BrokenRow {
    fn schema() -> &'static [ExcelColumn] {
        &[]
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self)
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Err(ExcelError::Format("row conversion failed".to_owned()))
    }
}

struct WideColumn;

impl ExcelRow for WideColumn {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("wide", "Wide", Some(65_536), 0, None)];
        COLUMNS
    }

    fn from_row(_row: &easyexcel_core::RowData) -> Result<Self> {
        Ok(Self)
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![CellValue::String("wide".to_owned())])
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
        },
        [every_cell()],
    )?;

    let mut workbook: Xlsx<_> = open_workbook(&path).map_err(format_error)?;
    let range = workbook.worksheet_range("Values").map_err(format_error)?;
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
        },
        [every_cell(), every_cell()],
    )?;
    let mut workbook: Xlsx<_> = open_workbook(path).map_err(format_error)?;
    let range = workbook.worksheet_range("Stream").map_err(format_error)?;
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
fn conversion_configuration_column_and_save_failures_propagate() -> Result<()> {
    let directory = tempdir()?;
    assert!(
        write_xlsx::<BrokenRow, _>(
            &directory.path().join("broken.xlsx"),
            &WriteOptions::default(),
            [BrokenRow]
        )
        .is_err()
    );
    assert!(
        write_xlsx::<WideColumn, _>(
            &directory.path().join("wide-head.xlsx"),
            &WriteOptions::default(),
            std::iter::empty()
        )
        .is_err()
    );
    assert!(
        write_xlsx::<WideColumn, _>(
            &directory.path().join("wide-data.xlsx"),
            &WriteOptions {
                need_head: false,
                ..WriteOptions::default()
            },
            [WideColumn]
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
            std::iter::empty()
        )
        .is_err()
    );
    assert!(
        write_xlsx::<EveryCell, _>(
            directory.path(),
            &WriteOptions::default(),
            std::iter::empty()
        )
        .is_err()
    );
    Ok(())
}
