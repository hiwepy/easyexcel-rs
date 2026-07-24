//! Mirrors Java `com.alibaba.excel.converters.date.DateStringConverter`.
//!
//! Rust maps Java `java.util.Date` to [`crate::JavaDate`].

/// Mirrors Java `DateStringConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct DateStringConverter;

impl crate::Converter<crate::JavaDate> for DateStringConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::String
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<crate::JavaDate, crate::ExcelError> {
        crate::converter::date_support::read_datetime(context).map(crate::JavaDate::new)
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, crate::JavaDate>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::date_support::write_datetime_string(
            context.value().naive_local(),
            context,
        ))
    }
}
