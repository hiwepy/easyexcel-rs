//! Mirrors Java `com.alibaba.excel.metadata.property.ExcelContentProperty`.

use crate::excel_cell_style::ExcelCellStyle;
use crate::excel_font_style::ExcelFontStyle;

/// Mirrors Java `ExcelContentProperty`.
///
/// Java carries a `Field`, `Converter`, `DateTimeFormatProperty`,
/// `NumberFormatProperty`, `StyleProperty`, `FontProperty`. Rust
/// collapses the format/style/font into the existing `ExcelCellStyle` /
/// `ExcelFontStyle` types and drops the reflection fields because the
/// derive macro handles them at compile time.
#[derive(Debug, Clone, Default)]
pub struct ExcelContentProperty {
    /// Content cell style. (Java `contentStyleProperty`)
    pub content_style_property: Option<ExcelCellStyle>,
    /// Content font style. (Java `contentFontProperty`)
    pub content_font_property: Option<ExcelFontStyle>,
    /// Optional date-time format string. (Java `dateTimeFormatProperty.format`)
    pub date_time_format: Option<&'static str>,
    /// Optional number format string. (Java `numberFormatProperty.format`)
    pub number_format: Option<&'static str>,
}

impl ExcelContentProperty {
    /// Creates an empty property. (Java `EMPTY = new ExcelContentProperty()`)
    pub const EMPTY: Self = Self {
        content_style_property: None,
        content_font_property: None,
        date_time_format: None,
        number_format: None,
    };
}
