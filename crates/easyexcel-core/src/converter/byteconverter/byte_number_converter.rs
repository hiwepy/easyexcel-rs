//! Mirrors Java `com.alibaba.excel.converters.byteconverter.ByteNumberConverter`.
//!
/// Mirrors Java `ByteNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct ByteNumberConverter;

impl crate::Converter<i8> for ByteNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<i8, crate::ExcelError> {
        crate::converter::number_support::read_number(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, i8>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        crate::converter::number_support::write_number(context)
    }
}
