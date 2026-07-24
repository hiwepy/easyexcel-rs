//! Mirrors Java `com.alibaba.excel.converters.longconverter.LongStringConverter`.
//!
/// Mirrors Java `LongStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LongStringConverter;

impl crate::Converter<i64> for LongStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<i64> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i64>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
