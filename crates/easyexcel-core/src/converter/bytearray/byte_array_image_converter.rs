//! Mirrors Java `com.alibaba.excel.converters.bytearray.ByteArrayImageConverter`.
//!
/// Mirrors Java `ByteArrayImageConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct ByteArrayImageConverter;

impl crate::Converter<Vec<u8>> for ByteArrayImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, Vec<u8>>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::WriteCellData::from_image(context.value().clone()))
    }
}
