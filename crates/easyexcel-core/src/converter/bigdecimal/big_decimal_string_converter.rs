//! Mirrors Java `com.alibaba.excel.converters.bigdecimal.BigDecimalStringConverter`.
//!
/// Mirrors Java `BigDecimalStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BigDecimalStringConverter;

impl crate::Converter<bigdecimal::BigDecimal> for BigDecimalStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<bigdecimal::BigDecimal> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, bigdecimal::BigDecimal>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
