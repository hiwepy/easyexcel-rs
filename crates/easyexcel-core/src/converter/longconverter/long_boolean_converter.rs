//! Mirrors Java `com.alibaba.excel.converters.longconverter.LongBooleanConverter`.
//!
/// Mirrors Java `LongBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LongBooleanConverter;

impl crate::Converter<i64> for LongBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<i64, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i64>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
