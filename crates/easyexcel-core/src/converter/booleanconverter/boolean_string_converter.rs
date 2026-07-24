//! Mirrors Java `com.alibaba.excel.converters.booleanconverter.BooleanStringConverter`.
//!
/// Mirrors Java `BooleanStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BooleanStringConverter;

impl crate::Converter<bool> for BooleanStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<bool, crate::ExcelError> {
        crate::converter::boolean_support::read_string_boolean(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, bool>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_boolean_string(
            context,
        ))
    }
}
