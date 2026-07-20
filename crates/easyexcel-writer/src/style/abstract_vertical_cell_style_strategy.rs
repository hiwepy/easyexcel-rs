//! Mirrors Java `com.alibaba.excel.write.style.AbstractVerticalCellStyleStrategy`.

use easyexcel_core::{ExcelCellStyle, WriteCellContext};

use crate::style::abstract_cell_style_strategy::AbstractCellStyleStrategy;

/// Mirrors Java `AbstractVerticalCellStyleStrategy extends AbstractCellStyleStrategy`.
///
/// The Java side stores two `WriteCellStyle` fields (`headCellStyle`,
/// `contentCellStyle`) and applies them based on `isHead`. The Rust
/// port exposes the same two accessors; concrete types such as
/// [`crate::style::vertical_cell_style_strategy::VerticalCellStyleStrategy`]
/// implement them and register as [`easyexcel_core::WriteHandler`].
///
/// Default methods return an empty style (Java returns `null`), so a
/// minimal override only fills the columns that need differentiation.
pub trait AbstractVerticalCellStyleStrategy: AbstractCellStyleStrategy {
    /// Returns the head cell style. (Java `headCellStyle(CellWriteHandlerContext)` /
    /// `headCellStyle(Head)`)
    fn head_cell_style(&self, _context: &WriteCellContext) -> ExcelCellStyle {
        ExcelCellStyle::new()
    }

    /// Returns the content cell style. (Java `contentCellStyle(CellWriteHandlerContext)` /
    /// `contentCellStyle(Head)`)
    fn content_cell_style(&self, _context: &WriteCellContext) -> ExcelCellStyle {
        ExcelCellStyle::new()
    }
}
