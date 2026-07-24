//! Mirrors Java `com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy`.

use easyexcel_core::WriteHandler;

/// Mirrors Java `AbstractColumnWidthStyleStrategy extends AbstractCellWriteHandler`.
///
/// Java declares a single `protected abstract Integer columnWidth(...)` hook.
pub trait AbstractColumnWidthStyleStrategy: WriteHandler {
    /// Returns the column width for the given column index, or `None` to keep
    /// the existing width. (Java `columnWidth(Head, Integer)`)
    fn column_width(&self, column_index: usize) -> Option<u16>;
}
