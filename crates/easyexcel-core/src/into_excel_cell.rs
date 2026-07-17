//! Mirrors the `convertToExcelData` half of Java `Converter<T>`.

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;

/// Converts a Rust value into a backend-neutral cell.
///
/// Java-side counterpart: `Converter<T>.convertToExcelData(...)`.
pub trait IntoExcelCell {
    /// Performs the conversion.
    ///
    /// # Errors
    ///
    /// Returns an error when the Rust value cannot be represented as an Excel cell.
    fn to_excel_cell(&self, context: &ConvertContext) -> Result<CellValue, ExcelError>;
}
