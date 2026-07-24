//! Mirrors Java `com.alibaba.excel.converters.bytearray.BoxingByteArrayImageConverter`.
//!
/// Mirrors Java `BoxingByteArrayImageConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct BoxingByteArrayImageConverter;

impl crate::Converter<Box<[u8]>> for BoxingByteArrayImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, Box<[u8]>>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::WriteCellData::from_image(context.value().to_vec()))
    }
}
