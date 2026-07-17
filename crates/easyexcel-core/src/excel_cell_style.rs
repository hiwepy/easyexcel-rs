//! Mirrors Java `com.alibaba.excel.write.metadata.style.WriteCellStyle` (the
//! annotation-driven subset carried by `ExcelCellStyle` / `ExcelFontStyle`).

use crate::excel_border_style::ExcelBorderStyle;
use crate::excel_color::ExcelColor;
use crate::excel_data_format::ExcelDataFormat;
use crate::excel_fill_pattern::ExcelFillPattern;
use crate::excel_horizontal_alignment::ExcelHorizontalAlignment;
use crate::excel_vertical_alignment::ExcelVerticalAlignment;

/// Cell-style properties generated from `HeadStyle` or `ContentStyle` equivalents.
///
/// All 23 fields correspond one-for-one to Java's `WriteCellStyle`. Java's
/// boxed `Short` / `Integer` becomes `Option<u16>` / `Option<i16>`; Java's
/// `BooleanEnum` becomes `Option<bool>`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ExcelCellStyle {
    /// Whether the cell is hidden when the sheet is protected.
    pub hidden: Option<bool>,
    /// Whether the cell is locked when the sheet is protected.
    pub locked: Option<bool>,
    /// Whether Excel treats the value as explicitly quoted text.
    pub quote_prefix: Option<bool>,
    /// Horizontal alignment.
    pub horizontal_alignment: Option<ExcelHorizontalAlignment>,
    /// Whether text wraps within the cell.
    pub wrapped: Option<bool>,
    /// Vertical alignment.
    pub vertical_alignment: Option<ExcelVerticalAlignment>,
    /// Text rotation in degrees.
    pub rotation: Option<i16>,
    /// Text indentation level.
    pub indent: Option<u8>,
    /// Left border style.
    pub border_left: Option<ExcelBorderStyle>,
    /// Right border style.
    pub border_right: Option<ExcelBorderStyle>,
    /// Top border style.
    pub border_top: Option<ExcelBorderStyle>,
    /// Bottom border style.
    pub border_bottom: Option<ExcelBorderStyle>,
    /// Left border indexed or RGB color.
    pub left_border_color: Option<ExcelColor>,
    /// Right border indexed or RGB color.
    pub right_border_color: Option<ExcelColor>,
    /// Top border indexed or RGB color.
    pub top_border_color: Option<ExcelColor>,
    /// Bottom border indexed or RGB color.
    pub bottom_border_color: Option<ExcelColor>,
    /// Fill pattern.
    pub fill_pattern: Option<ExcelFillPattern>,
    /// Fill background indexed or RGB color.
    pub fill_background_color: Option<ExcelColor>,
    /// Fill foreground indexed or RGB color.
    pub fill_foreground_color: Option<ExcelColor>,
    /// Whether text shrinks to fit the cell.
    pub shrink_to_fit: Option<bool>,
    /// Built-in or custom Excel number format.
    pub data_format: Option<ExcelDataFormat>,
}

impl ExcelCellStyle {
    /// Creates an annotation style with every property unspecified. (Java `WriteCellStyle()`)
    #[must_use]
    pub const fn new() -> Self {
        Self {
            hidden: None,
            locked: None,
            quote_prefix: None,
            horizontal_alignment: None,
            wrapped: None,
            vertical_alignment: None,
            rotation: None,
            indent: None,
            border_left: None,
            border_right: None,
            border_top: None,
            border_bottom: None,
            left_border_color: None,
            right_border_color: None,
            top_border_color: None,
            bottom_border_color: None,
            fill_pattern: None,
            fill_background_color: None,
            fill_foreground_color: None,
            shrink_to_fit: None,
            data_format: None,
        }
    }
}
