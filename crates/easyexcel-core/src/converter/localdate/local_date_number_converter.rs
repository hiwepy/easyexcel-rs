//! Mirrors Java `com.alibaba.excel.converters.localdate.LocalDateNumberConverter`.
//!
/// Mirrors Java `LocalDateNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct LocalDateNumberConverter;

impl crate::Converter<chrono::NaiveDate> for LocalDateNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
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
        Ok(crate::WriteCellData::new(crate::CellValue::Float(
            crate::converter::date_support::date_to_excel_serial(
                *context.value(),
                context.convert_context().use_1904_windowing,
            ),
        )))
    }
}
