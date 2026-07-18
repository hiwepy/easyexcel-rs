//! Mirrors Java `com.alibaba.excel.write.metadata.WriteWorkbook`.

use easyexcel_core::CsvCharset;

use crate::WriteOptions;

/// Mirrors Java `WriteWorkbook extends WriteBasicParameter`.
///
/// The Java side carries 11 fields (file, outputStream, templateFile, etc.).
/// Rust reuses the existing [`WriteOptions`] struct that already models the
/// same data; this newtype exists so the public API carries a 1:1 named
/// class and lets builders accept either `WriteOptions` or `WriteWorkbook`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteWorkbook {
    /// Backing configuration. (Java `WriteWorkbook` getter surface)
    pub options: WriteOptions,
}

impl WriteWorkbook {
    /// Creates a new `WriteWorkbook` with default options.
    #[must_use]
    pub fn new() -> Self {
        Self {
            options: WriteOptions::default(),
        }
    }

    /// Returns the effective write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }

    /// Returns the charset. (Java `getCharset()`)
    #[must_use]
    pub fn charset(&self) -> &CsvCharset {
        &self.options.charset
    }

    /// Returns the password, if any. (Java `getPassword()`)
    #[must_use]
    pub fn password(&self) -> Option<&str> {
        self.options.password.as_deref()
    }

    /// Returns the in-memory flag. (Java `getInMemory()`)
    #[must_use]
    pub const fn in_memory(&self) -> bool {
        !self.options.constant_memory
    }

    /// Returns the write-on-exception flag. (Java `getWriteExcelOnException()`)
    #[must_use]
    pub const fn write_excel_on_exception(&self) -> bool {
        self.options.write_excel_on_exception
    }
}

impl Default for WriteWorkbook {
    fn default() -> Self {
        Self::new()
    }
}

impl From<WriteOptions> for WriteWorkbook {
    fn from(options: WriteOptions) -> Self {
        Self { options }
    }
}