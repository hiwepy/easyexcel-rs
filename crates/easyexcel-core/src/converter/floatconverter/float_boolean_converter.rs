//! Mirrors Java `com.alibaba.excel.converters.floatconverter.FloatBooleanConverter`.
//!
/// Mirrors Java `FloatBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct FloatBooleanConverter;

impl crate::Converter<f32> for FloatBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<f32, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, f32>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
