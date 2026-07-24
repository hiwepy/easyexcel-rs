//! Mirrors Java `com.alibaba.excel.converters.biginteger.BigIntegerNumberConverter`.
//!
/// Mirrors Java `BigIntegerNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BigIntegerNumberConverter;

impl crate::Converter<num_bigint::BigInt> for BigIntegerNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<num_bigint::BigInt, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, num_bigint::BigInt>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
