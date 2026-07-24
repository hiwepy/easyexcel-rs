//! Mirrors Java `com.alibaba.excel.converters.shortconverter.ShortBooleanConverter`.
//!
/// Mirrors Java `ShortBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct ShortBooleanConverter;

impl crate::Converter<i16> for ShortBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<i16, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i16>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
