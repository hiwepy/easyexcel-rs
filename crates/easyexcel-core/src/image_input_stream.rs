//! Mirrors Java `com.alibaba.excel.converters.inputstream.InputStreamImageConverter`.

use std::cell::RefCell;
use std::io::Read;

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;
use crate::into_excel_cell::IntoExcelCell;

/// Java `InputStreamImageConverter` equivalent for a stateful Rust [`Read`] source.
///
/// Conversion consumes the bytes remaining in the reader and deliberately
/// does not close or replace it, matching Java `EasyExcel`'s ownership
/// contract for a caller-supplied `InputStream`.
#[derive(Debug)]
pub struct ImageInputStream<R> {
    reader: RefCell<R>,
}

impl<R> ImageInputStream<R> {
    /// Wraps a reader whose remaining bytes represent one image.
    #[must_use]
    pub const fn new(reader: R) -> Self {
        Self {
            reader: RefCell::new(reader),
        }
    }

    /// Returns the wrapped reader, preserving its position after conversion.
    #[must_use]
    pub fn into_inner(self) -> R {
        self.reader.into_inner()
    }
}

impl<R> From<R> for ImageInputStream<R> {
    fn from(reader: R) -> Self {
        Self::new(reader)
    }
}

impl<R: Read> IntoExcelCell for ImageInputStream<R> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        read_image_bytes(&mut *self.reader.borrow_mut()).map(CellValue::Image)
    }
}

fn read_image_bytes(reader: &mut dyn Read) -> Result<Vec<u8>, ExcelError> {
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    Ok(bytes)
}
