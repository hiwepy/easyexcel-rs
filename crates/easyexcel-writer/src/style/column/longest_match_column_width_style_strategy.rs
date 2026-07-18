//! Mirrors Java `com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy`.

use easyexcel_core::WriteHandler;

use crate::style::column::abstract_head_column_width_style_strategy::AbstractHeadColumnWidthStyleStrategy;

/// Mirrors Java `LongestMatchColumnWidthStyleStrategy`.
///
/// Java walks the rendered sheet after write to derive the longest
/// content width per column. The Rust port uses
/// `worksheet.autofit()` on save to achieve the same effect.
pub struct LongestMatchColumnWidthStyleStrategy;

impl LongestMatchColumnWidthStyleStrategy {
    /// Creates the strategy.
    #[must_use]
    pub const fn new() -> Self {
        Self
    }
}

impl Default for LongestMatchColumnWidthStyleStrategy {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for LongestMatchColumnWidthStyleStrategy {
    fn order(&self) -> i32 {
        // Mirror Java `OrderConstant.DEFINE_STYLE` to let autofit run late.
        -50_000
    }
}

impl AbstractHeadColumnWidthStyleStrategy for LongestMatchColumnWidthStyleStrategy {
    fn head_column_width(&self, _column_index: usize) -> Option<u16> {
        // `rust_xlsxwriter` autofit reads the actual cell content widths.
        None
    }
}