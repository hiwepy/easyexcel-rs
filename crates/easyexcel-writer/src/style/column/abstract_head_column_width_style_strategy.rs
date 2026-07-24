//! Mirrors Java `com.alibaba.excel.write.style.column.AbstractHeadColumnWidthStyleStrategy`.

use easyexcel_core::WriteHandler;

/// Mirrors Java `AbstractHeadColumnWidthStyleStrategy`.
pub trait AbstractHeadColumnWidthStyleStrategy: WriteHandler {
    /// Returns the head column width. (Java `columnWidth(Head, Integer)`)
    fn head_column_width(&self, column_index: usize) -> Option<u16>;
}
