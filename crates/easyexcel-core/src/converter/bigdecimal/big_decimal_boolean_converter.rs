//! Mirrors Java `com.alibaba.excel.converters.bigdecimal.BigDecimalBooleanConverter`.
//!
/// Mirrors Java `BigDecimalBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BigDecimalBooleanConverter;

impl crate::Converter<bigdecimal::BigDecimal> for BigDecimalBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<bigdecimal::BigDecimal, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, bigdecimal::BigDecimal>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
