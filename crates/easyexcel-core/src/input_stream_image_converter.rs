//! Mirrors Java `com.alibaba.excel.converters.inputstream.InputStreamImageConverter`
//! (sentinel type).

use std::io::Read;

use crate::cell_value::CellValue;
use crate::converter::Converter;
use crate::excel_error::ExcelError;
use crate::image_input_stream::ImageInputStream;
use crate::into_excel_cell::IntoExcelCell;
use crate::write_cell_data::WriteCellData;
use crate::write_converter_context::WriteConverterContext;

/// Java `InputStreamImageConverter` equivalent for annotation-selected stream fields.
#[derive(Debug, Clone, Copy, Default)]
pub struct InputStreamImageConverter;

impl<R: Read> Converter<ImageInputStream<R>> for InputStreamImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, ImageInputStream<R>>,
    ) -> Result<WriteCellData, ExcelError> {
        let value = context.value().to_excel_cell(context.convert_context())?;
        match value {
            CellValue::Image(bytes) => Ok(WriteCellData::from_image(bytes)),
            other => Ok(WriteCellData::new(other)),
        }
    }
}
