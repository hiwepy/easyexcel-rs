//! Mirrors Java `com.alibaba.excel.write.style.AbstractVerticalCellStyleStrategy`.

use easyexcel_core::{ExcelCellStyle, WriteCellContext};

use crate::style::abstract_cell_style_strategy::AbstractCellStyleStrategy;

/// Mirrors Java `AbstractVerticalCellStyleStrategy extends AbstractCellStyleStrategy`.
///
/// The Java side stores two `WriteCellStyle` fields (`headCellStyle`,
/// `contentCellStyle`) and applies them based on `isHead`. The Rust
/// port exposes the same two accessors and `WriteHandler` hooks
/// dispatch through them.
pub trait AbstractVerticalCellStyleStrategy: AbstractCellStyleStrategy {
    /// Returns the head cell style. (Java `getHeadCellStyle(CellWriteHandlerContext)`)
    fn head_cell_style(&self, _context: &WriteCellContext) -> ExcelCellStyle;

    /// Returns the content cell style. (Java `getContentCellStyle(CellWriteHandlerContext)`)
    fn content_cell_style(&self, _context: &WriteCellContext) -> ExcelCellStyle;
}