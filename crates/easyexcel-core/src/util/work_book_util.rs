//! Mirrors Java com.alibaba.excel.util.WorkBookUtil.
//!
//! Java wraps Apache POI `Workbook` / `Sheet` / `Row` / `Cell`
//! construction behind a small utility boundary. Rust keeps that boundary
//! backend-neutral: writer crates implement the creator traits for XLSX,
//! BIFF8 or CSV objects and these functions perform the same delegation.

use crate::excel_error::ExcelError;
use crate::write_cell_data::WriteCellData;

/// Backend factory used by [`create_work_book`].
pub trait WorkBookCreator {
    /// Concrete workbook produced by this backend.
    type WorkBook;

    /// Creates or opens the workbook.
    ///
    /// # Errors
    ///
    /// Returns an I/O, format or unsupported-operation error from the backend.
    fn create_work_book(self) -> Result<Self::WorkBook, ExcelError>;
}

/// Backend workbook capable of creating a sheet.
pub trait SheetCreator {
    /// Concrete sheet handle returned by this backend.
    type Sheet<'a>
    where
        Self: 'a;

    /// Creates a sheet with the supplied name.
    ///
    /// # Errors
    ///
    /// Returns a format error for an invalid or duplicate sheet name.
    fn create_sheet(&mut self, sheet_name: &str) -> Result<Self::Sheet<'_>, ExcelError>;
}

/// Backend sheet capable of creating a logical row.
pub trait RowCreator {
    /// Concrete row handle returned by this backend.
    type Row<'a>
    where
        Self: 'a;

    /// Creates a row at a zero-based index.
    ///
    /// # Errors
    ///
    /// Returns a format error when the row is outside the backend limit.
    fn create_row(&mut self, row_index: u32) -> Result<Self::Row<'_>, ExcelError>;
}

/// Backend row capable of creating a logical cell.
pub trait CellCreator {
    /// Concrete cell handle returned by this backend.
    type Cell<'a>
    where
        Self: 'a;

    /// Creates a cell at a zero-based column index.
    ///
    /// # Errors
    ///
    /// Returns a format error when the column is outside the backend limit.
    fn create_cell(&mut self, column_index: u16) -> Result<Self::Cell<'_>, ExcelError>;
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createWorkBook`.
///
/// # Errors
///
/// Propagates workbook construction errors from the selected backend.
pub fn create_work_book<C: WorkBookCreator>(creator: C) -> Result<C::WorkBook, ExcelError> {
    creator.create_work_book()
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createSheet`.
///
/// # Errors
///
/// Propagates sheet creation errors from the selected backend.
pub fn create_sheet<'a, C: SheetCreator>(
    workbook: &'a mut C,
    sheet_name: &str,
) -> Result<C::Sheet<'a>, ExcelError> {
    workbook.create_sheet(sheet_name)
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createRow`.
///
/// # Errors
///
/// Propagates row creation errors from the selected backend.
pub fn create_row<C: RowCreator>(sheet: &mut C, row_index: u32) -> Result<C::Row<'_>, ExcelError> {
    sheet.create_row(row_index)
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createCell`.
///
/// # Errors
///
/// Propagates cell creation errors from the selected backend.
pub fn create_cell<C: CellCreator>(
    row: &mut C,
    column_index: u16,
) -> Result<C::Cell<'_>, ExcelError> {
    row.create_cell(column_index)
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#fillDataFormat`.
///
/// Java creates the missing `WriteCellStyle` and `DataFormatData` containers,
/// then sets the requested format only when no format was already assigned.
pub fn fill_data_format(cell_data: &mut WriteCellData, format: Option<&str>, default_format: &str) {
    cell_data.get_or_create_style();
    let data_format = cell_data.get_or_create_data_format();
    if data_format.format.is_none() {
        data_format.format = Some(format.unwrap_or(default_format).to_owned());
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellValue, ExcelCellStyle};

    #[derive(Default)]
    struct TestWorkBook {
        sheets: Vec<TestSheet>,
    }

    struct TestWorkBookFactory;

    impl WorkBookCreator for TestWorkBookFactory {
        type WorkBook = TestWorkBook;

        fn create_work_book(self) -> Result<Self::WorkBook, ExcelError> {
            Ok(TestWorkBook::default())
        }
    }

    #[derive(Default)]
    struct TestSheet {
        name: String,
        rows: Vec<TestRow>,
    }

    #[derive(Default)]
    struct TestRow {
        index: u32,
        cells: Vec<TestCell>,
    }

    #[derive(Debug, PartialEq, Eq)]
    struct TestCell {
        column: u16,
    }

    impl SheetCreator for TestWorkBook {
        type Sheet<'a> = &'a mut TestSheet;

        fn create_sheet(&mut self, sheet_name: &str) -> Result<Self::Sheet<'_>, ExcelError> {
            self.sheets.push(TestSheet {
                name: sheet_name.to_owned(),
                rows: Vec::new(),
            });
            Ok(self.sheets.last_mut().expect("just pushed"))
        }
    }

    impl RowCreator for TestSheet {
        type Row<'a> = &'a mut TestRow;

        fn create_row(&mut self, row_index: u32) -> Result<Self::Row<'_>, ExcelError> {
            self.rows.push(TestRow {
                index: row_index,
                cells: Vec::new(),
            });
            Ok(self.rows.last_mut().expect("just pushed"))
        }
    }

    impl CellCreator for TestRow {
        type Cell<'a> = &'a mut TestCell;

        fn create_cell(&mut self, column_index: u16) -> Result<Self::Cell<'_>, ExcelError> {
            self.cells.push(TestCell {
                column: column_index,
            });
            Ok(self.cells.last_mut().expect("just pushed"))
        }
    }

    #[test]
    fn creator_chain_delegates_to_real_backend_objects() {
        let mut workbook = create_work_book(TestWorkBookFactory).expect("workbook");
        let sheet = create_sheet(&mut workbook, "用户").expect("sheet");
        assert_eq!(sheet.name, "用户");
        let row = create_row(sheet, 7).expect("row");
        assert_eq!(row.index, 7);
        let cell = create_cell(row, 3).expect("cell");
        assert_eq!(*cell, TestCell { column: 3 });
        assert_eq!(workbook.sheets[0].rows[0].cells.len(), 1);
    }

    #[test]
    fn fill_data_format_creates_nested_state_and_preserves_existing_format() {
        let mut cell = WriteCellData::new(CellValue::Int(1));
        fill_data_format(&mut cell, None, "yyyy-mm-dd");
        assert_eq!(
            cell.data_format_data().and_then(|value| value.format()),
            Some("yyyy-mm-dd")
        );
        assert_eq!(cell.write_cell_style(), Some(&ExcelCellStyle::default()));

        fill_data_format(&mut cell, Some("0.00"), "General");
        assert_eq!(
            cell.data_format_data().and_then(|value| value.format()),
            Some("yyyy-mm-dd"),
            "Java does not overwrite an existing format"
        );
    }
}
