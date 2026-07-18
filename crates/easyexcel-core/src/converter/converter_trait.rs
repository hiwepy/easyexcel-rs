//! Mirrors `com.alibaba.excel.converters.Converter<T>` public surface.
//!
//! Java exposes four default methods:
//! * `convertToJavaData(ReadCellData, ExcelContentProperty, GlobalConfiguration)`
//! * `convertToJavaData(ReadConverterContext)`
//! * `convertToExcelData(T, ExcelContentProperty, GlobalConfiguration)`
//! * `convertToExcelData(WriteConverterContext)`
//!
//! plus `supportJavaTypeKey` / `supportExcelTypeKey`.
//!
//! Rust keeps `support_excel_type` (used as the read dispatch key together
//! with `TypeId`) and the two conversion methods. `supportJavaTypeKey` is
//! implicit in the generic parameter `T`.

use crate::cell_value::CellValue;
use crate::enum_cell_data_type::CellDataType;
use crate::excel_error::ExcelError;
use crate::read_converter_context::ReadConverterContext;
use crate::write_converter_context::WriteConverterContext;

/// Custom bidirectional converter selected by `#[excel(converter = Type)]`.
///
/// The Java counterpart exposes six default methods (`supportJavaTypeKey`,
/// `supportExcelTypeKey`, two `convertToJavaData` overloads, two
/// `convertToExcelData` overloads). Rust's idiomatic trait surface keeps
/// `support_excel_type` (read dispatch key) plus the two conversion methods;
/// `supportJavaTypeKey` is encoded by the generic parameter `T` and the
/// `ConverterRegistry::register::<T, _>` call.
#[allow(clippy::missing_errors_doc)]
pub trait Converter<T> {
    /// Returns the source cell type supported when this converter is registered globally.
    ///
    /// Java `EasyExcel` requires global read converters to expose this key.
    /// A string default keeps field-only converters concise while matching
    /// the most common custom converter contract.
    fn support_excel_type(&self) -> CellDataType {
        CellDataType::String
    }

    /// Converts an Excel cell into a Rust field value. (Java `convertToJavaData(ReadConverterContext)`)
    fn convert_to_rust_data(&self, _context: &ReadConverterContext<'_>) -> Result<T, ExcelError> {
        Err(ExcelError::Unsupported(
            "custom converter does not support reading".to_owned(),
        ))
    }

    /// Converts a Rust field value into an Excel cell. (Java `convertToExcelData(WriteConverterContext)`)
    fn convert_to_excel_data(
        &self,
        _context: &WriteConverterContext<'_, T>,
    ) -> Result<CellValue, ExcelError> {
        Err(ExcelError::Unsupported(
            "custom converter does not support writing".to_owned(),
        ))
    }
}
