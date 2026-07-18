//! Mirrors Java com.alibaba.excel.util.IoUtils.

#![allow(dead_code)]

use std::io::{self, Read, Write};

use crate::excel_error::ExcelError;

/// Mirrors `org.apache.commons.io.IOUtils#copy`.
///
/// Copies all bytes from `reader` into `writer` using a 4 KiB stack
/// buffer (Java uses a 4 KiB byte array).
pub fn copy(reader: &mut dyn Read, writer: &mut dyn Write) -> Result<u64, ExcelError> {
    let n = io::copy(reader, writer)?;
    Ok(n)
}
