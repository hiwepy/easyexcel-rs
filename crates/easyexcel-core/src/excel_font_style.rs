//! Mirrors Java `com.alibaba.excel.write.metadata.style.WriteFont` (the
//! annotation-driven subset carried by `ExcelFontStyle`).

use crate::excel_color::ExcelColor;
use crate::excel_font_script::ExcelFontScript;
use crate::excel_underline::ExcelUnderline;

/// Font properties generated from `HeadFontStyle` or `ContentFontStyle` equivalents.
///
/// All nine fields correspond one-for-one to Java's `WriteFont`. `font_name`
/// is constrained to `&'static str` so the struct can stay `Copy`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ExcelFontStyle {
    /// Font family name.
    pub font_name: Option<&'static str>,
    /// Font size in points.
    pub font_height_in_points: Option<f64>,
    /// Italic rendering.
    pub italic: Option<bool>,
    /// Strike-through rendering.
    pub strikeout: Option<bool>,
    /// Font indexed or RGB color.
    pub color: Option<ExcelColor>,
    /// Superscript or subscript positioning.
    pub type_offset: Option<ExcelFontScript>,
    /// Underline rendering.
    pub underline: Option<ExcelUnderline>,
    /// Font character set.
    pub charset: Option<u8>,
    /// Bold rendering.
    pub bold: Option<bool>,
}

impl ExcelFontStyle {
    /// Creates an annotation font style with every property unspecified. (Java `WriteFont()`)
    #[must_use]
    pub const fn new() -> Self {
        Self {
            font_name: None,
            font_height_in_points: None,
            italic: None,
            strikeout: None,
            color: None,
            type_offset: None,
            underline: None,
            charset: None,
            bold: None,
        }
    }
}
