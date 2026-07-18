//! Mirrors Java `com.alibaba.excel.write.handler.CellWriteHandler`.

use easyexcel_core::{WriteCellContext, WriteRowContext};

/// Mirrors Java `CellWriteHandler extends WriteHandler`.
///
/// Rust collapses the four capability interfaces
/// (`Workbook/Sheet/Row/CellWriteHandler`) into a single trait. This
/// module re-exports the cell-level surface for parity with the Java
/// package layout.
pub trait CellWriteHandler: easyexcel_core::WriteHandler {
    /// Called before a cell is written. (Java `beforeCellCreate(CellWriteHandlerContext)`)
    fn before_cell_create(&mut self, _context: &mut WriteCellContext) {}

    /// Called after the cell data is converted. (Java `afterCellDataConverted`)
    fn after_cell_data_converted(&mut self, _context: &WriteCellContext) {}

    /// Called after a cell has been processed. (Java `afterCellDispose`)
    fn after_cell_dispose(&mut self, _context: &WriteCellContext) {}

    /// Convenience: row context access. (Java `afterRowDispose` row signal)
    fn after_row_dispose_marker(&self) -> Option<&WriteRowContext> {
        None
    }
}