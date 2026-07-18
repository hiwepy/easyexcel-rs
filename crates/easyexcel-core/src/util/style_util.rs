//! Mirrors Java com.alibaba.excel.util.StyleUtil.
//!
//! Java wraps Apache POI `CellStyle`, `RichTextString`,
//! `HyperlinkType`, `Font`, and `DataFormat` helpers. The Rust port
//! delegates the same operations to `rust_xlsxwriter`, which provides
//! them through its own API (`Format`, `Format::set_bold`, etc.).
//! These functions are kept as minimal placeholders so the 1:1 Java
//! file mapping is preserved.

#![allow(dead_code)]

use std::any::Any;

/// Mirrors `com.alibaba.excel.util.StyleUtil#buildCellStyle`.
#[must_use]
pub fn build_cell_style() -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.StyleUtil#buildRichTextString`.
#[must_use]
pub fn build_rich_text_string(_text: &str) -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.StyleUtil#getCellCoordinate`.
#[must_use]
pub fn get_cell_coordinate(_rich_text_string: &dyn Any) -> Option<String> {
    None
}

/// Mirrors `com.alibaba.excel.util.StyleUtil#getHyperlinkType`.
#[must_use]
pub fn get_hyperlink_type(_hyperlink: &dyn Any) -> &'static str {
    "none"
}

/// Mirrors `com.alibaba.excel.util.StyleUtil#buildFont`.
#[must_use]
pub fn build_font() -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.StyleUtil#buildDataFormat`.
#[must_use]
pub fn build_data_format(_format: &str) -> Option<Box<dyn Any>> {
    None
}
