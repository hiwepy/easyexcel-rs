//! Mirrors Java `com.alibaba.excel.write.metadata.style.WriteFont`.

use crate::excel_color::ExcelColor;
use crate::excel_font_script::ExcelFontScript;
use crate::excel_underline::ExcelUnderline;

/// Runtime font metadata equivalent to Java `WriteFont`.
///
/// Java uses boxed `Boolean`/`Short`/`Byte`/`Integer`; Rust uses `Option`
/// to express "unset" with zero overhead. All nine fields preserve the Java
/// semantics, including the POI alignment with `null` meaning "inherit".
#[derive(Debug, Clone, Default, PartialEq)]
pub struct WriteFont {
    font_name: Option<String>,
    font_height_in_points: Option<f64>,
    italic: Option<bool>,
    strikeout: Option<bool>,
    color: Option<ExcelColor>,
    type_offset: Option<ExcelFontScript>,
    underline: Option<ExcelUnderline>,
    charset: Option<u8>,
    bold: Option<bool>,
}

impl WriteFont {
    /// Creates font metadata with every property unspecified.
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

    /// Sets the font family name. (Java `setFontName(String)`)
    #[must_use]
    pub fn font_name(mut self, value: impl Into<String>) -> Self {
        self.font_name = Some(value.into());
        self
    }

    /// Sets the font size in points. (Java `setFontHeightInPoints(Short)`)
    #[must_use]
    pub const fn font_height_in_points(mut self, value: f64) -> Self {
        self.font_height_in_points = Some(value);
        self
    }

    /// Sets italic rendering. (Java `setItalic(Boolean)`)
    #[must_use]
    pub const fn italic(mut self, value: bool) -> Self {
        self.italic = Some(value);
        self
    }

    /// Sets strike-through rendering. (Java `setStrikeout(Boolean)`)
    #[must_use]
    pub const fn strikeout(mut self, value: bool) -> Self {
        self.strikeout = Some(value);
        self
    }

    /// Sets an indexed or RGB font color. (Java `setColor(Short)`)
    #[must_use]
    pub const fn color(mut self, value: ExcelColor) -> Self {
        self.color = Some(value);
        self
    }

    /// Sets superscript or subscript rendering. (Java `setTypeOffset(Short)`)
    #[must_use]
    pub const fn type_offset(mut self, value: ExcelFontScript) -> Self {
        self.type_offset = Some(value);
        self
    }

    /// Sets underline rendering. (Java `setUnderline(Byte)`)
    #[must_use]
    pub const fn underline(mut self, value: ExcelUnderline) -> Self {
        self.underline = Some(value);
        self
    }

    /// Sets the font character set. (Java `setCharset(Integer)`)
    #[must_use]
    pub const fn charset(mut self, value: u8) -> Self {
        self.charset = Some(value);
        self
    }

    /// Sets bold rendering. (Java `setBold(Boolean)`)
    #[must_use]
    pub const fn bold(mut self, value: bool) -> Self {
        self.bold = Some(value);
        self
    }

    /// Returns the optional font family name. (Java `getFontName()`)
    #[must_use]
    pub fn get_font_name(&self) -> Option<&str> {
        self.font_name.as_deref()
    }

    /// Returns the optional font size. (Java `getFontHeightInPoints()`)
    #[must_use]
    pub const fn get_font_height_in_points(&self) -> Option<f64> {
        self.font_height_in_points
    }

    /// Returns the optional italic flag. (Java `getItalic()`)
    #[must_use]
    pub const fn get_italic(&self) -> Option<bool> {
        self.italic
    }

    /// Returns the optional strike-through flag. (Java `getStrikeout()`)
    #[must_use]
    pub const fn get_strikeout(&self) -> Option<bool> {
        self.strikeout
    }

    /// Returns the optional font color. (Java `getColor()`)
    #[must_use]
    pub const fn get_color(&self) -> Option<ExcelColor> {
        self.color
    }

    /// Returns the optional superscript/subscript mode. (Java `getTypeOffset()`)
    #[must_use]
    pub const fn get_type_offset(&self) -> Option<ExcelFontScript> {
        self.type_offset
    }

    /// Returns the optional underline mode. (Java `getUnderline()`)
    #[must_use]
    pub const fn get_underline(&self) -> Option<ExcelUnderline> {
        self.underline
    }

    /// Returns the optional character set. (Java `getCharset()`)
    #[must_use]
    pub const fn get_charset(&self) -> Option<u8> {
        self.charset
    }

    /// Returns the optional bold flag. (Java `getBold()`)
    #[must_use]
    pub const fn get_bold(&self) -> Option<bool> {
        self.bold
    }
}
