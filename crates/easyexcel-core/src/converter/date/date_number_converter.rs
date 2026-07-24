//! Mirrors Java `com.alibaba.excel.converters.date.DateNumberConverter`.
//!
//! Rust maps Java `java.util.Date` to [`crate::JavaDate`].

/// Mirrors Java `DateNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct DateNumberConverter;

impl crate::Converter<crate::JavaDate> for DateNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
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
        Ok(crate::WriteCellData::new(crate::CellValue::Float(
            crate::converter::date_support::datetime_to_excel_serial(
                context.value().naive_local(),
                context.convert_context().use_1904_windowing,
            ),
        )))
    }
}
