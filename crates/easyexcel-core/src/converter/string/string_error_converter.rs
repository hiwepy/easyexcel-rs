//! Mirrors Java `com.alibaba.excel.converters.string.StringErrorConverter`.
//!
/// Mirrors Java `StringErrorConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct StringErrorConverter;

impl crate::Converter<String> for StringErrorConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Error
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<String, crate::ExcelError> {
        match context.cell().unwrap_or(&crate::CellValue::Empty) {
            crate::CellValue::Error(value) => Ok(value.clone()),
            value => Err(context.convert_context().invalid(value, "String")),
        }
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, String>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::WriteCellData::new(crate::CellValue::Error(
            context.value().clone(),
        )))
    }
}
