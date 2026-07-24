//! Mirrors Java `com.alibaba.excel.converters.doubleconverter.DoubleStringConverter`.
//!
/// Mirrors Java `DoubleStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct DoubleStringConverter;

impl crate::Converter<f64> for DoubleStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<f64> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, f64>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
