//! Mirrors Java `com.alibaba.excel.converters.string.StringStringConverter`.
//!
/// Mirrors Java `StringStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct StringStringConverter;

impl crate::Converter<String> for StringStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<String, crate::ExcelError> {
        match context.cell().unwrap_or(&crate::CellValue::Empty) {
            crate::CellValue::String(value) => Ok(value.clone()),
            value => Err(context.convert_context().invalid(value, "String")),
        }
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, String>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::WriteCellData::from_string(context.value().clone()))
    }
}
