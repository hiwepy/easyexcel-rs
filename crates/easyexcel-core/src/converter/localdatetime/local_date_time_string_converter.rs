//! Mirrors Java `com.alibaba.excel.converters.localdatetime.LocalDateTimeStringConverter`.
//!
/// Mirrors Java `LocalDateTimeStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LocalDateTimeStringConverter;

impl crate::Converter<chrono::NaiveDateTime> for LocalDateTimeStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
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
        Ok(crate::converter::date_support::write_datetime_string(
            *context.value(),
            context,
        ))
    }
}
