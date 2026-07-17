//! Mirrors Java `com.alibaba.excel.write.handler.WriteHandler` plus its four
//! capability interfaces (Workbook/Sheet/Row/CellWriteHandler) collapsed into a
//! single trait.

use crate::analysis_context::Result;
use crate::cell_value::CellValue;
use crate::write_cell_context::WriteCellContext;
use crate::write_row_context::WriteRowContext;
use crate::write_sheet_context::WriteSheetContext;
use crate::write_workbook_context::WriteWorkbookContext;

/// Intercepts the workbook, worksheet, row, and cell write lifecycle.
///
/// Java exposes four capability interfaces (`WorkbookWriteHandler`,
/// `SheetWriteHandler`, `RowWriteHandler`, `CellWriteHandler`) plus four
/// `Abstract*WriteHandler` skeletons. Rust collapses the eight lifecycle
/// hooks into a single trait; default implementations are no-ops so a
/// minimal handler only overrides the events it cares about.
///
/// Any callback may return an error to stop the write immediately.
#[allow(clippy::missing_errors_doc)]
pub trait WriteHandler {
    /// Lower orders execute first. (Java `Order.order()`)
    fn order(&self) -> i32 {
        0
    }

    /// Called before the workbook is created. (Java `WorkbookWriteHandler.beforeWorkbookCreate`)
    fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Called after the workbook is saved. (Java `WorkbookWriteHandler.afterWorkbookDispose`)
    fn after_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Called after the worksheet is configured and before rows are written.
    /// (Java `SheetWriteHandler.beforeSheetCreate`)
    fn before_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Called after all worksheet rows are written. (Java `SheetWriteHandler.afterSheetCreate`)
    fn after_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Called before a row is written. (Java `RowWriteHandler.beforeRowCreate`)
    fn before_row(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Called after a row is written. (Java `RowWriteHandler.afterRowCreate`)
    fn after_row(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Called before a cell is written and may transform or skip it.
    /// (Java `CellWriteHandler.beforeCellCreate`)
    fn before_cell(&mut self, _context: &mut WriteCellContext) -> Result<()> {
        Ok(())
    }

    /// Called after a cell has been processed. (Java `CellWriteHandler.afterCellDispose`)
    fn after_cell(&mut self, _context: &WriteCellContext) -> Result<()> {
        Ok(())
    }
}

// `CellValue` import retained for downstream conversions; suppress unused warning.
#[allow(dead_code)]
fn _import_marker(_: CellValue) {}
