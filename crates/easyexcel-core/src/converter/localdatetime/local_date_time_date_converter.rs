//! Mirrors Java `com.alibaba.excel.converters.localdatetime.LocalDateTimeDateConverter`.
//!
/// Mirrors Java `LocalDateTimeDateConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LocalDateTimeDateConverter;

impl crate::Converter<chrono::NaiveDateTime> for LocalDateTimeDateConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Date
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, chrono::NaiveDateTime>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::date_support::write_datetime_value(
            *context.value(),
            context,
        ))
    }
}
