//! Mirrors Java `com.alibaba.excel.converters.date.DateDateConverter`.
//!
//! Rust maps Java `java.util.Date` to [`crate::JavaDate`].

/// Mirrors Java `DateDateConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct DateDateConverter;

impl crate::Converter<crate::JavaDate> for DateDateConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Date
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, crate::JavaDate>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        Ok(crate::converter::date_support::write_datetime_value(
            context.value().naive_local(),
            context,
        ))
    }
}
