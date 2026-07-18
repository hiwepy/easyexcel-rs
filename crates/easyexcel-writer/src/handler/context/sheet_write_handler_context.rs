//! Mirrors Java `com.alibaba.excel.write.handler.context.SheetWriteHandlerContext`.

use easyexcel_core::WriteSheetContext;

/// Mirrors Java `SheetWriteHandlerContext`.
pub struct SheetWriteHandlerContext {
    /// Mirrors the sheet context.
    pub sheet: WriteSheetContext,
}

impl SheetWriteHandlerContext {
    /// Returns the sheet context. (Java `getSheet()` step)
    #[must_use]
    pub const fn sheet(&self) -> &WriteSheetContext {
        &self.sheet
    }
}