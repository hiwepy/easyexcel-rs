//! Mirrors Java `com.alibaba.excel.write.handler.context.RowWriteHandlerContext`.

use easyexcel_core::WriteRowContext;

/// Mirrors Java `RowWriteHandlerContext`.
pub struct RowWriteHandlerContext {
    /// Mirrors the row context.
    pub row: WriteRowContext,
}

impl RowWriteHandlerContext {
    /// Returns the row context. (Java `getRow()` step)
    #[must_use]
    pub const fn row(&self) -> &WriteRowContext {
        &self.row
    }
}