//! Mirrors Java `com.alibaba.excel.write.style.row.AbstractRowHeightStyleStrategy`.

use easyexcel_core::WriteHandler;

/// Mirrors Java `AbstractRowHeightStyleStrategy`.
pub trait AbstractRowHeightStyleStrategy: WriteHandler {
    /// Returns the head row height, or `None` for default.
    fn head_row_height(&self) -> Option<u16>;

    /// Returns the content row height, or `None` for default.
    fn content_row_height(&self) -> Option<u16>;
}