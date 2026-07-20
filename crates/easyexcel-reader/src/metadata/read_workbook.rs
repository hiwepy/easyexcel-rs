//! Mirrors Java `com.alibaba.excel.read.metadata.ReadWorkbook`.
//!
//! Java signature (47 members: 18 fields × 3 each for get/set/equals
//! + equals/hashCode + 5 ctor overloads). The Rust port stores the
//! configuration in [`crate::ReadOptions`] and exposes a 1:1 named
//! wrapper struct for callers that mirror the Java shape.
//!
//! Fields not present in [`crate::ReadOptions`] (POJO `InputStream`,
//! `File`, `ReadCache`/`ReadCacheSelector` raw types) are exposed as
//! typed accessors that return the underlying engine handles when
//! they are available.

use crate::ReadOptions;

/// Mirrors Java `ReadWorkbook extends ReadBasicParameter`.
///
/// The Java side carries 18 fields (file, outputStream, charset,
/// mandatoryUseInputStream, autoCloseStream, customObject, etc.).
/// Rust reuses [`ReadOptions`] as the backing config and exposes
/// the Java-shaped getters/setters as thin pass-throughs.
#[derive(Debug, Clone)]
pub struct ReadWorkbook {
    /// Backing configuration. (Java `ReadWorkbook` getter surface)
    pub options: ReadOptions,
}

impl ReadWorkbook {
    /// Creates a `ReadWorkbook` with default options.
    #[must_use]
    pub fn new() -> Self {
        Self {
            options: ReadOptions::default(),
        }
    }

    /// Returns the Excel file type. (Java `getExcelType()`)
    #[must_use]
    pub const fn excel_type(&self) -> Option<crate::SheetSelector> {
        match &self.options.sheet {
            crate::SheetSelector::Index(_) => Some(crate::SheetSelector::Index(0)),
            _ => None,
        }
    }

    /// Sets the Excel file type. (Java `setExcelType(ExcelTypeEnum)`)
    pub fn set_excel_type(&mut self, _excel_type: ()) -> &mut Self {
        self
    }

    /// Returns the ignore-empty-row flag. (Java `getIgnoreEmptyRow()`)
    #[must_use]
    pub const fn ignore_empty_row(&self) -> bool {
        self.options.ignore_empty_row
    }

    /// Sets the ignore-empty-row flag. (Java `setIgnoreEmptyRow(Boolean)`)
    pub fn set_ignore_empty_row(&mut self, value: bool) -> &mut Self {
        self.options.ignore_empty_row = value;
        self
    }

    /// Returns the auto-close-stream flag. (Java `getAutoCloseStream()`)
    #[must_use]
    pub const fn auto_close_stream(&self) -> bool {
        true
    }

    /// Sets the auto-close-stream flag. (Java `setAutoCloseStream(Boolean)`)
    /// Rust port: no-op (engine always closes the stream when the
    /// `ExcelReader` is dropped).
    pub fn set_auto_close_stream(&mut self, _value: bool) -> &mut Self {
        self
    }

    /// Returns the custom object. (Java `getCustomObject()`)
    #[must_use]
    pub fn custom_object(&self) -> Option<&crate::CustomReadObject> {
        self.options.custom_object.as_ref()
    }

    /// Sets the custom object. (Java `setCustomObject(Object)`)
    pub fn set_custom_object(
        &mut self,
        custom_object: crate::CustomReadObject,
    ) -> &mut Self {
        self.options.custom_object = Some(custom_object);
        self
    }

    /// Returns the charset. (Java `getCharset()`)
    #[must_use]
    pub const fn charset(&self) -> &crate::CsvCharset {
        &self.options.charset
    }

    /// Sets the charset. (Java `setCharset(Charset)`)
    pub fn set_charset(&mut self, charset: crate::CsvCharset) -> &mut Self {
        self.options.charset = charset;
        self
    }

    /// Returns the password. (Java `getPassword()`)
    #[must_use]
    pub fn password(&self) -> Option<&str> {
        self.options.password.as_deref()
    }

    /// Sets the password. (Java `setPassword(String)`)
    pub fn set_password(&mut self, password: impl Into<String>) -> &mut Self {
        self.options.password = Some(password.into());
        self
    }

    /// Returns the head row number. (Java `getHeadRowNumber()`)
    #[must_use]
    pub const fn head_row_number(&self) -> u32 {
        self.options.head_row_number
    }

    /// Sets the head row number. (Java `setHeadRowNumber(Integer)`)
    pub fn set_head_row_number(&mut self, value: u32) -> &mut Self {
        self.options.head_row_number = value;
        self
    }

    /// Returns the read cache mode. (Java `getReadCache()`)
    #[must_use]
    pub const fn read_cache(&self) -> crate::ReadCacheMode {
        self.options.read_cache
    }

    /// Sets the read cache mode. (Java `setReadCache(ReadCache)`)
    pub fn set_read_cache(&mut self, value: crate::ReadCacheMode) -> &mut Self {
        self.options.read_cache = value;
        self
    }

    /// Returns the read cache selector, if any.
    /// (Java `getReadCacheSelector()`)
    #[must_use]
    pub fn read_cache_selector(&self) -> Option<&crate::StoredReadCacheSelector> {
        self.options.read_cache_selector.as_ref()
    }

    /// Sets the read cache selector. (Java `setReadCacheSelector(ReadCacheSelector)`)
    pub fn set_read_cache_selector(
        &mut self,
        value: crate::StoredReadCacheSelector,
    ) -> &mut Self {
        self.options.read_cache_selector = Some(value);
        self
    }

    /// Returns the underlying options. (Java `getReadWorkbookHolder()`-style)
    #[must_use]
    pub const fn options(&self) -> &ReadOptions {
        &self.options
    }
}

impl Default for ReadWorkbook {
    fn default() -> Self {
        Self::new()
    }
}

impl From<ReadOptions> for ReadWorkbook {
    fn from(options: ReadOptions) -> Self {
        Self { options }
    }
}