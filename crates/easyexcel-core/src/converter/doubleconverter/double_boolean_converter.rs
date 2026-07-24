//! Mirrors Java `com.alibaba.excel.converters.doubleconverter.DoubleBooleanConverter`.
//!
/// Mirrors Java `DoubleBooleanConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct DoubleBooleanConverter;

impl crate::Converter<f64> for DoubleBooleanConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Boolean
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<f64, crate::ExcelError> {
        crate::converter::boolean_support::read_boolean_scalar(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, f64>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::boolean_support::write_scalar_boolean(
            context,
        ))
    }
}
