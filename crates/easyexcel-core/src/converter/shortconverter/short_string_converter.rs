//! Mirrors Java `com.alibaba.excel.converters.shortconverter.ShortStringConverter`.
//!
/// Mirrors Java `ShortStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct ShortStringConverter;

impl crate::Converter<i16> for ShortStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<i16> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i16>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
