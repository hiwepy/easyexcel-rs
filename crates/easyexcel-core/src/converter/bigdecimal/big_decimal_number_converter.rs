//! Mirrors Java `com.alibaba.excel.converters.bigdecimal.BigDecimalNumberConverter`.
//!
/// Mirrors Java `BigDecimalNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BigDecimalNumberConverter;

impl crate::Converter<bigdecimal::BigDecimal> for BigDecimalNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<bigdecimal::BigDecimal, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, bigdecimal::BigDecimal>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
