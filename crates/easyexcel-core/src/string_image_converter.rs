//! Mirrors Java `com.alibaba.excel.converters.string.StringImageConverter`.
//!
//! Used with `#[excel(converter = StringImageConverter)]`. The file path is
//! read during row conversion; missing or unreadable files return an I/O
//! error.

use crate::cell_value::CellValue;
use crate::converter::Converter;
use crate::excel_error::ExcelError;
use crate::write_converter_context::WriteConverterContext;

/// Java `StringImageConverter` equivalent for fields containing an image file path.
///
/// Use it with `#[excel(converter = StringImageConverter)]`. The file is
/// read during row conversion; missing or unreadable files return an I/O
/// error.
#[derive(Debug, Clone, Copy, Default)]
pub struct StringImageConverter;

impl Converter<String> for StringImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<CellValue, ExcelError> {
        std::fs::read(context.value())
            .map(CellValue::Image)
            .map_err(Into::into)
    }
}
