//! Mirrors Java `com.alibaba.excel.metadata.property.ExcelContentProperty` and
//! `com.alibaba.excel.metadata.GlobalConfiguration` (subset).

use crate::excel_error::ExcelError;

/// Location and formatting information supplied to cell converters.
///
/// Java's `ReadConverterContext` and `WriteConverterContext` carry
/// `contentProperty` (resolved annotation) plus `analysisContext` or
/// `writeContext`. Rust collapses them into a single `Copy` value so each
/// cell conversion can pass it by reference without ownership fuss.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConvertContext {
    /// Sheet name. (Java `AnalysisContext.readSheetHolder().getSheetName()`)
    pub sheet_name: String,
    /// Zero-based row index. (Java `AnalysisContext.readRowHolder().getRowIndex()`)
    pub row_index: u32,
    /// Zero-based column index when it can be resolved.
    pub column_index: Option<usize>,
    /// Rust field name. (Java `ExcelContentProperty.getField().getName()`)
    pub field: &'static str,
    /// Optional format string. (Java `ExcelContentProperty.getDateTimeFormatProperty()` etc.)
    pub format: Option<&'static str>,
    /// Whether numeric dates use Excel's 1904 date system.
    /// (Java `GlobalConfiguration.getUse1904windowing()`)
    pub use_1904_windowing: bool,
}

impl ConvertContext {
    /// Builds a typed conversion error matching Java `ExcelDataConvertException`.
    pub(crate) fn invalid(
        &self,
        value: &crate::cell_value::CellValue,
        target: &'static str,
    ) -> ExcelError {
        ExcelError::Data {
            sheet: self.sheet_name.clone(),
            row: self.row_index,
            column: self.column_index,
            field: self.field,
            value: value.as_text(),
            message: format!("cannot convert cell to {target}"),
        }
    }

    /// Attaches this field's conversion location to an arbitrary converter error.
    ///
    /// Java wraps every exception raised by `Converter.convertToExcelData` in
    /// `ExcelWriteDataConvertException`, so even converters that return a
    /// generic error retain the field, row, and column from the active cell
    /// handler context. Derive-generated Rust writers use this helper before
    /// the backend replaces the provisional row/column with their physical
    /// worksheet coordinates.
    #[must_use]
    pub fn write_error(&self, error: ExcelError) -> ExcelError {
        let (value, message) = match error {
            ExcelError::Data { value, message, .. } => (value, message),
            other => (String::new(), other.to_string()),
        };
        ExcelError::Data {
            sheet: self.sheet_name.clone(),
            row: self.row_index,
            column: self.column_index,
            field: self.field,
            value,
            message,
        }
    }
}
