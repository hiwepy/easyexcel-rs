//! Mirrors Java `com.alibaba.excel.converters.shortconverter.ShortNumberConverter`.
//!
/// Mirrors Java `ShortNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct ShortNumberConverter;

impl crate::Converter<i16> for ShortNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<i16, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i16>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
