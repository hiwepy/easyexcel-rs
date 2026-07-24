//! Mirrors Java `com.alibaba.excel.converters.doubleconverter.DoubleNumberConverter`.
//!
/// Mirrors Java `DoubleNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct DoubleNumberConverter;

impl crate::Converter<f64> for DoubleNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<f64, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, f64>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
