//! Mirrors Java `com.alibaba.excel.write.handler.RowWriteHandler`.

use easyexcel_core::WriteRowContext;

/// Mirrors Java `RowWriteHandler extends WriteHandler`.
pub trait RowWriteHandler: easyexcel_core::WriteHandler {
    /// Called before a row is created. (Java `beforeRowCreate`)
    fn before_row_create(&mut self, _context: &WriteRowContext) {}

    /// Called after a row is created. (Java `afterRowCreate`)
    fn after_row_create(&mut self, _context: &WriteRowContext) {}

    /// Called after the row has been processed. (Java `afterRowDispose`)
    fn after_row_dispose(&mut self, _context: &WriteRowContext) {}
}