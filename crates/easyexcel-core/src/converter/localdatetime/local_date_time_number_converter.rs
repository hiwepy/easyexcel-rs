//! Mirrors Java `com.alibaba.excel.converters.localdatetime.LocalDateTimeNumberConverter`.
//!
/// Mirrors Java `LocalDateTimeNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LocalDateTimeNumberConverter;

impl crate::Converter<chrono::NaiveDateTime> for LocalDateTimeNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<chrono::NaiveDateTime, crate::ExcelError> {
        crate::converter::date_support::read_datetime(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, chrono::NaiveDateTime>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::WriteCellData::new(crate::CellValue::Float(
            crate::converter::date_support::datetime_to_excel_serial(
                *context.value(),
                context.convert_context().use_1904_windowing,
            ),
        )))
    }
}
