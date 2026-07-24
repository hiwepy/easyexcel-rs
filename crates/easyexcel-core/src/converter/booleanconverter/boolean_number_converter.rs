//! Mirrors Java `com.alibaba.excel.converters.booleanconverter.BooleanNumberConverter`.
//!
/// Mirrors Java `BooleanNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BooleanNumberConverter;

impl crate::Converter<bool> for BooleanNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<bool, crate::ExcelError> {
        crate::converter::boolean_support::read_number_boolean(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, bool>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_boolean_number(
            context,
        ))
    }
}
