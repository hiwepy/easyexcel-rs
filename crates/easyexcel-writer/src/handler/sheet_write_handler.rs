//! Mirrors Java `com.alibaba.excel.write.handler.SheetWriteHandler`.

use easyexcel_core::WriteSheetContext;

/// Mirrors Java `SheetWriteHandler extends WriteHandler`.
pub trait SheetWriteHandler: easyexcel_core::WriteHandler {
    /// Called before a sheet is created. (Java `beforeSheetCreate`)
    fn before_sheet_create(&mut self, _context: &WriteSheetContext) {}

    /// Called after a sheet is created. (Java `afterSheetCreate`)
    fn after_sheet_create(&mut self, _context: &WriteSheetContext) {}
}