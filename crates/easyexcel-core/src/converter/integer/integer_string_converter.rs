//! Mirrors Java `com.alibaba.excel.converters.integer.IntegerStringConverter`.
//!
/// Mirrors Java `IntegerStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct IntegerStringConverter;

impl crate::Converter<i32> for IntegerStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<i32> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i32>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
