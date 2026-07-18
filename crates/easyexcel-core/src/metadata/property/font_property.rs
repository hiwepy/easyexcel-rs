//! Mirrors Java `com.alibaba.excel.metadata.property.FontProperty`.

use crate::excel_color::ExcelColor;
use crate::excel_font_script::ExcelFontScript;
use crate::excel_underline::ExcelUnderline;

/// Mirrors Java `FontProperty`. Rust reuses `ExcelFontStyle` for the
/// runtime representation; this struct exists for 1:1 Java package
/// parity.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FontProperty {
    /// Font family name. (Java `fontName`)
    pub font_name: Option<&'static str>,
    /// Font size in points. (Java `fontHeightInPoints`)
    pub font_height_in_points: Option<f64>,
    /// Italic. (Java `italic`)
    pub italic: Option<bool>,
    /// Strike-through. (Java `strikeout`)
    pub strikeout: Option<bool>,
    /// Color. (Java `color`)
    pub color: Option<ExcelColor>,
    /// Super/subscript. (Java `typeOffset`)
    pub type_offset: Option<ExcelFontScript>,
    /// Underline. (Java `underline`)
    pub underline: Option<ExcelUnderline>,
    /// Character set. (Java `charset`)
    pub charset: Option<u8>,
    /// Bold. (Java `bold`)
    pub bold: Option<bool>,
}
