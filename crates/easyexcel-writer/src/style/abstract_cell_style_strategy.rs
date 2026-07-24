//! Mirrors Java `com.alibaba.excel.write.style.AbstractCellStyleStrategy`.

use easyexcel_core::{WriteCellContext, WriteHandler};

/// Mirrors Java `AbstractCellStyleStrategy extends AbstractCellWriteHandler`.
///
/// The Java side selects a `WriteCellStyle` for each cell. Rust collapses
/// the strategy into a [`WriteHandler`] trait implementation that
/// returns a style from a per-call closure.
pub trait AbstractCellStyleStrategy: WriteHandler {
    /// Returns the cell style to apply for the current cell.
    fn cell_style(&self, context: &WriteCellContext) -> easyexcel_core::ExcelCellStyle;
}
