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
    /// Mirrors `WriteWorkbook.excelType`. (Java `getExcelType()`)
    pub excel_type: easyexcel_core::support::ExcelTypeEnum,
}

impl WriteWorkbook {
    /// Creates a new `WriteWorkbook` with default options.
    #[must_use]
    pub fn new() -> Self {
        Self {
            options: WriteOptions::default(),
            excel_type: easyexcel_core::support::ExcelTypeEnum::Xlsx,
        }
    }

    /// Returns the effective write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }

    /// Returns the Excel file type. (Java `getExcelType()`)
    #[must_use]
    pub fn excel_type(&self) -> easyexcel_core::support::ExcelTypeEnum {
        self.excel_type
    }

    /// Sets the Excel file type. (Java `setExcelType(ExcelTypeEnum)`)
    pub fn set_excel_type(
        &mut self,
        excel_type: easyexcel_core::support::ExcelTypeEnum,
    ) -> &mut Self {
        self.excel_type = excel_type;
        self
    }

    /// Returns the output file path. (Java `getFile()`)
    ///
    /// Rust port: file is a constructor concern on `ExcelWriter`
    /// rather than a `WriteOptions` field. The getter returns the
    /// sheet name as a 1:1 placeholder so callers can mirror the
    /// Java shape; the actual path lives on `ExcelWriter::output_path()`.
    #[must_use]
    pub fn file(&self) -> Option<&std::path::Path> {
        None
    }

    /// Sets the output file path. (Java `setFile(File)`)
    ///
    /// Rust port: no-op; configure file via
    /// `EasyExcel::write(path).build()`.
    pub fn set_file(&mut self, _file: impl Into<std::path::PathBuf>) -> &mut Self {
        self
    }

    /// Returns the charset. (Java `getCharset()`)
    #[must_use]
    pub fn charset(&self) -> &CsvCharset {
        &self.options.charset
    }

    /// Sets the charset. (Java `setCharset(Charset)`)
    pub fn set_charset(&mut self, charset: CsvCharset) -> &mut Self {
        self.options.charset = charset;
        self
    }

    /// Returns the BOM flag. (Java `getWithBom()`)
    #[must_use]
    pub const fn with_bom(&self) -> bool {
        self.options.with_bom
    }

    /// Sets the BOM flag. (Java `setWithBom(Boolean)`)
    pub fn set_with_bom(&mut self, with_bom: bool) -> &mut Self {
        self.options.with_bom = with_bom;
        self
    }

    /// Returns the password, if any. (Java `getPassword()`)
    #[must_use]
    pub fn password(&self) -> Option<&str> {
        self.options.password.as_deref()
    }

    /// Sets the password. (Java `setPassword(String)`)
    pub fn set_password(&mut self, password: impl Into<String>) -> &mut Self {
        self.options.password = Some(password.into());
        self
    }

    /// Returns the in-memory flag. (Java `getInMemory()`)
    #[must_use]
    pub const fn in_memory(&self) -> bool {
        !self.options.constant_memory
    }

    /// Sets the in-memory flag. (Java `setInMemory(boolean)`)
    pub fn set_in_memory(&mut self, in_memory: bool) -> &mut Self {
        self.options.constant_memory = !in_memory;
        self
    }

    /// Returns the write-on-exception flag. (Java `getWriteExcelOnException()`)
    #[must_use]
    pub const fn write_excel_on_exception(&self) -> bool {
        self.options.write_excel_on_exception
    }

    /// Sets the write-on-exception flag. (Java `setWriteExcelOnException(boolean)`)
    pub fn set_write_excel_on_exception(&mut self, value: bool) -> &mut Self {
        self.options.write_excel_on_exception = value;
        self
    }

    /// Returns the auto-close-stream flag. (Java `getAutoCloseStream()`)
    #[must_use]
    pub const fn auto_close_stream(&self) -> bool {
        self.options.auto_close_stream
    }

    /// Sets the auto-close-stream flag. (Java `setAutoCloseStream(boolean)`)
    pub fn set_auto_close_stream(&mut self, value: bool) -> &mut Self {
        self.options.auto_close_stream = value;
        self
    }
}

impl Default for WriteWorkbook {
    fn default() -> Self {
        Self::new()
    }
}

impl From<WriteOptions> for WriteWorkbook {
    fn from(options: WriteOptions) -> Self {
        Self {
            options,
            excel_type: easyexcel_core::support::ExcelTypeEnum::Xlsx,
        }
    }
}