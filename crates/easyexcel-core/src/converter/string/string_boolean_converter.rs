//! Mirrors Java `com.alibaba.excel.converters.string.StringBooleanConverter`.
//!
/// Mirrors Java `StringBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct StringBooleanConverter;

impl crate::Converter<String> for StringBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<String, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_string_value(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, String>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_string_boolean(
            context,
        ))
    }
}
