//! Mirrors Java `com.alibaba.excel.metadata.property.StyleProperty`.

use crate::excel_cell_style::ExcelCellStyle;

/// Mirrors Java `StyleProperty`. Rust reuses `ExcelCellStyle` for the
/// runtime representation; this struct exists for 1:1 Java package
/// parity.
/// `Eq` is not derived because [`ExcelCellStyle`] embeds `f64` font size.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct StyleProperty {
    /// The underlying cell style. (Java delegates all fields)
    pub cell_style: ExcelCellStyle,
}

impl StyleProperty {
    /// Creates a `StyleProperty`. (Java constructor)
    #[must_use]
    pub const fn new(cell_style: ExcelCellStyle) -> Self {
        Self { cell_style }
    }
}
