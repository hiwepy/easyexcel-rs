//! Mirrors Java `com.alibaba.excel.converters.biginteger.BigIntegerBooleanConverter`.
//!
/// Mirrors Java `BigIntegerBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BigIntegerBooleanConverter;

impl crate::Converter<num_bigint::BigInt> for BigIntegerBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<num_bigint::BigInt, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, num_bigint::BigInt>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
