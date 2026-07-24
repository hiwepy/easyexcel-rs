//! Mirrors Java `com.alibaba.excel.converters.biginteger.BigIntegerStringConverter`.
//!
/// Mirrors Java `BigIntegerStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BigIntegerStringConverter;

impl crate::Converter<num_bigint::BigInt> for BigIntegerStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }
    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> crate::Result<num_bigint::BigInt> {
        crate::converter::number_support::read_string_number(context)
    }
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, num_bigint::BigInt>,
    ) -> crate::Result<crate::WriteCellData> {
        crate::converter::number_support::write_number_string(context)
    }
}
