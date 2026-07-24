//! Mirrors Java `com.alibaba.excel.converters.inputstream.InputStreamImageConverter`.

use std::cell::RefCell;
use std::fmt;
use std::io::Read;

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;
use crate::from_excel_cell::FromExcelCell;
use crate::into_excel_cell::IntoExcelCell;

/// Java `InputStreamImageConverter` equivalent for a stateful Rust [`Read`] source.
///
/// The first conversion consumes and caches the bytes remaining in the reader;
/// repeated conversion passes reuse that cache. The reader is deliberately not
/// closed or replaced, matching Java `EasyExcel`'s ownership contract for a
/// caller-supplied `InputStream`.
pub struct ImageInputStream<R = Box<dyn Read + Send>> {
    reader: RefCell<R>,
    cached_bytes: RefCell<Option<Vec<u8>>>,
}

impl<R> fmt::Debug for ImageInputStream<R> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ImageInputStream")
            .finish_non_exhaustive()
    }
}

impl<R> ImageInputStream<R> {
    /// Wraps a reader whose remaining bytes represent one image.
    #[must_use]
    pub const fn new(reader: R) -> Self {
        Self {
            reader: RefCell::new(reader),
            cached_bytes: RefCell::new(None),
        }
    }

    /// Returns the wrapped reader, preserving its position after conversion.
    #[must_use]
    pub fn into_inner(self) -> R {
        self.reader.into_inner()
    }
}

impl ImageInputStream {
    /// Type-erases a reader so the default converter registry can use one stable `TypeId`.
    ///
    /// This is the Rust counterpart of declaring a Java model field as
    /// `InputStream` rather than as a concrete `ByteArrayInputStream` subtype.
    #[must_use]
    pub fn boxed<R>(reader: R) -> Self
    where
        R: Read + Send + 'static,
    {
        Self::new(Box::new(reader))
    }
}

impl<R> From<R> for ImageInputStream<R> {
    fn from(reader: R) -> Self {
        Self::new(reader)
    }
}

impl<R> FromExcelCell for ImageInputStream<R> {
    fn from_excel_cell(
        _value: Option<&CellValue>,
        _context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        Err(ExcelError::Unsupported(
            "InputStreamImageConverter does not support reading image cells".to_owned(),
        ))
    }
}

impl<R: Read> IntoExcelCell for ImageInputStream<R> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        if let Some(bytes) = self.cached_bytes.borrow().as_ref() {
            return Ok(CellValue::Image(bytes.clone()));
        }
        let bytes = read_image_bytes(&mut *self.reader.borrow_mut())?;
        *self.cached_bytes.borrow_mut() = Some(bytes.clone());
        Ok(CellValue::Image(bytes))
    }
}

fn read_image_bytes(reader: &mut dyn Read) -> Result<Vec<u8>, ExcelError> {
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    Ok(bytes)
}
