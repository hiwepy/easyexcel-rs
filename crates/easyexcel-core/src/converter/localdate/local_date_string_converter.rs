//! Mirrors Java `com.alibaba.excel.converters.localdate.LocalDateStringConverter`.
//!
/// Mirrors Java `LocalDateStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LocalDateStringConverter;

impl crate::Converter<chrono::NaiveDate> for LocalDateStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<chrono::NaiveDate, crate::ExcelError> {
        crate::converter::date_support::read_date(context)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, chrono::NaiveDate>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::date_support::write_date_string(
            *context.value(),
            context,
        ))
    }
}
