//! Mirrors the `convertToJavaData` half of Java `Converter<T>`.

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;

/// Converts a backend-neutral cell into a Rust value.
///
/// Java-side counterpart: `Converter<T>.convertToJavaData(...)`.
/// `Sized` bound matches Java's ability to instantiate the result type.
pub trait FromExcelCell: Sized {
    /// Performs the conversion.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Data`] when the cell cannot be represented by `Self`.
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError>;
}
