//! Mirrors Java `com.alibaba.excel.converters.byteconverter.ByteBooleanConverter`.
//!
/// Mirrors Java `ByteBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct ByteBooleanConverter;

impl crate::Converter<i8> for ByteBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<i8, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i8>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
