//! Mirrors Java `com.alibaba.excel.write.handler.context.WorkbookWriteHandlerContext`.

use easyexcel_core::WriteWorkbookContext;

/// Mirrors Java `WorkbookWriteHandlerContext`.
pub struct WorkbookWriteHandlerContext {
    /// Mirrors the workbook context.
    pub workbook: WriteWorkbookContext,
}

impl WorkbookWriteHandlerContext {
    /// Returns the workbook context. (Java `getWorkbook()` step)
    #[must_use]
    pub const fn workbook(&self) -> &WriteWorkbookContext {
        &self.workbook
    }
}