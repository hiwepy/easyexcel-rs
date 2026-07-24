//! Mirrors Java `com.alibaba.excel.converters.localdate.LocalDateDateConverter`.
//!
/// Mirrors Java `LocalDateDateConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LocalDateDateConverter;

impl crate::Converter<chrono::NaiveDate> for LocalDateDateConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Date
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, chrono::NaiveDate>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::date_support::write_date_value(
            *context.value(),
            context,
        ))
    }
}
