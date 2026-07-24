//! Mirrors Java `com.alibaba.excel.converters.integer.IntegerNumberConverter`.
//!
/// Mirrors Java `IntegerNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct IntegerNumberConverter;

impl crate::Converter<i32> for IntegerNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<i32, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i32>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
