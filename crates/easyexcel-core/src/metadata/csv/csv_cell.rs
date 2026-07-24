//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvCell`.

use crate::{CellValue, NumericCellType};

use super::csv_cell_style::CsvCellStyle;
use super::csv_rich_text_string::CsvRichTextString;

/// A typed cell in the logical CSV workbook model.
#[derive(Debug, Clone, PartialEq)]
pub struct CsvCell {
    column_index: u16,
    value: CellValue,
    numeric_cell_type: Option<NumericCellType>,
    cell_style: Option<CsvCellStyle>,
}

impl CsvCell {
    /// Creates an empty cell at a zero-based column index.
    #[must_use]
    pub const fn new(column_index: u16) -> Self {
        Self {
            column_index,
            value: CellValue::Empty,
            numeric_cell_type: None,
            cell_style: None,
        }
    }

    /// Returns the zero-based column index.
    #[must_use]
    pub const fn column_index(&self) -> u16 {
        self.column_index
    }

    /// Returns the typed value.
    #[must_use]
    pub const fn value(&self) -> &CellValue {
        &self.value
    }

    /// Replaces the typed value.
    pub fn set_value(&mut self, value: impl Into<CellValue>) {
        self.value = value.into();
        self.numeric_cell_type = match self.value {
            CellValue::Date(_) | CellValue::DateTime(_) => Some(NumericCellType::Date),
            CellValue::Int(_) | CellValue::Float(_) | CellValue::Decimal(_) => {
                Some(NumericCellType::Number)
            }
            _ => None,
        };
    }

    /// Stores a formula value.
    pub fn set_formula(&mut self, formula: impl Into<String>) {
        self.value = CellValue::Formula(formula.into());
        self.numeric_cell_type = None;
    }

    /// Stores plain text from a CSV rich-text wrapper.
    pub fn set_rich_text(&mut self, value: CsvRichTextString) {
        self.value = CellValue::String(value.as_str().to_owned());
        self.numeric_cell_type = None;
    }

    /// Returns whether the numeric payload represents a date or a number.
    #[must_use]
    pub const fn numeric_cell_type(&self) -> Option<NumericCellType> {
        self.numeric_cell_type
    }

    /// Applies a CSV cell style.
    pub fn set_cell_style(&mut self, style: CsvCellStyle) {
        self.cell_style = Some(style);
    }

    /// Returns the applied CSV style.
    #[must_use]
    pub const fn cell_style(&self) -> Option<&CsvCellStyle> {
        self.cell_style.as_ref()
    }

    /// Returns the display value written to the CSV record.
    #[must_use]
    pub fn display_text(&self) -> String {
        self.value.as_text()
    }
}
