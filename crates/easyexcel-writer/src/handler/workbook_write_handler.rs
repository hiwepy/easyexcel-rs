//! Mirrors Java `com.alibaba.excel.write.handler.WorkbookWriteHandler`.

use easyexcel_core::WriteWorkbookContext;

/// Mirrors Java `WorkbookWriteHandler extends WriteHandler`.
pub trait WorkbookWriteHandler: easyexcel_core::WriteHandler {
    /// Called before a workbook is created. (Java `beforeWorkbookCreate`)
    fn before_workbook_create(&mut self, _context: &WriteWorkbookContext) {}

    /// Called after a workbook is created. (Java `afterWorkbookCreate`)
    fn after_workbook_create(&mut self, _context: &WriteWorkbookContext) {}

    /// Called after the workbook has been disposed. (Java `afterWorkbookDispose`)
    fn after_workbook_dispose(&mut self, _context: &WriteWorkbookContext) {}
}