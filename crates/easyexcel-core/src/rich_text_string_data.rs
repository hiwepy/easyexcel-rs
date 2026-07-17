//! Mirrors Java `com.alibaba.excel.metadata.data.RichTextStringData`.

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;
use crate::from_excel_cell::FromExcelCell;
use crate::interval_font::IntervalFont;
use crate::into_excel_cell::IntoExcelCell;
use crate::write_font::WriteFont;

/// Java `RichTextStringData` equivalent.
///
/// Java exposes `textString`, `writeFont`, `intervalFontList` via Lombok
/// accessors. Rust preserves the same fields and offers builder-style
/// `apply_font` / `apply_font_range` setters matching the Java semantics.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichTextStringData {
    text_string: String,
    write_font: Option<WriteFont>,
    interval_font_list: Vec<IntervalFont>,
}

impl RichTextStringData {
    /// Creates rich-text metadata for a string. (Java `RichTextStringData(String)`)
    #[must_use]
    pub fn new(text_string: impl Into<String>) -> Self {
        Self {
            text_string: text_string.into(),
            write_font: None,
            interval_font_list: Vec::new(),
        }
    }

    /// Applies a font to the entire string. (Java `applyFont(WriteFont)`)
    #[must_use]
    pub fn apply_font(mut self, write_font: WriteFont) -> Self {
        self.write_font = Some(write_font);
        self
    }

    /// Applies a font to a half-open UTF-16 character range. (Java `applyFont(int, int, WriteFont)`)
    #[must_use]
    pub fn apply_font_range(
        mut self,
        start_index: usize,
        end_index: usize,
        write_font: WriteFont,
    ) -> Self {
        self.interval_font_list
            .push(IntervalFont::new(start_index, end_index, write_font));
        self
    }

    /// Replaces all interval font entries.
    #[must_use]
    pub fn interval_font_list(mut self, value: impl IntoIterator<Item = IntervalFont>) -> Self {
        self.interval_font_list = value.into_iter().collect();
        self
    }

    /// Returns the underlying text. (Java `getTextString()`)
    #[must_use]
    pub fn text_string(&self) -> &str {
        &self.text_string
    }

    /// Returns the optional whole-string font. (Java `getWriteFont()`)
    #[must_use]
    pub const fn write_font(&self) -> Option<&WriteFont> {
        self.write_font.as_ref()
    }

    /// Returns interval fonts in application order. (Java `getIntervalFontList()`)
    #[must_use]
    pub fn interval_fonts(&self) -> &[IntervalFont] {
        &self.interval_font_list
    }
}

impl IntoExcelCell for RichTextStringData {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::RichText(self.clone()))
    }
}

impl FromExcelCell for RichTextStringData {
    fn from_excel_cell(
        cell: Option<&CellValue>,
        _context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        Ok(Self::new(cell.map_or_else(String::new, CellValue::as_text)))
    }
}
