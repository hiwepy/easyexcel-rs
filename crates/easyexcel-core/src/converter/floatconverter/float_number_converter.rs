//! Mirrors Java `com.alibaba.excel.converters.floatconverter.FloatNumberConverter`.
//!
/// Mirrors Java `FloatNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct FloatNumberConverter;

impl crate::Converter<f32> for FloatNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<f32, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, f32>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
