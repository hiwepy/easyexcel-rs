//! Mirrors Java `com.alibaba.excel.converters.booleanconverter.BooleanBooleanConverter`.
//!
/// Mirrors Java `BooleanBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BooleanBooleanConverter;

impl crate::Converter<bool> for BooleanBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<bool, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, bool>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::WriteCellData::new(crate::CellValue::Bool(
            *context.value(),
        )))
    }
}
