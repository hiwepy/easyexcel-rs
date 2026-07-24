//! Mirrors Java `com.alibaba.excel.converters.floatconverter.FloatStringConverter`.
//!
/// Mirrors Java `FloatStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct FloatStringConverter;

impl crate::Converter<f32> for FloatStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<f32> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, f32>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
