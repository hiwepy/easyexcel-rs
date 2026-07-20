//! Mirrors Java `com.alibaba.excel.metadata.BasicParameter`.

use crate::CacheLocation;

/// Shared read/write builder parameters.
///
/// Java stores a reflective `Class<?> clazz`; Rust stores the type name string
/// because model metadata is resolved at compile time through `ExcelRow`.
///
/// Rust port of Java `BasicParameter`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct BasicParameter {
    /// Dynamic header rows. (Java `head`)
    pub head: Option<Vec<Vec<String>>>,
    /// Model type name. (Java `clazz`)
    pub clazz: Option<String>,
    /// Custom converter type names registered on the builder. (Java `customConverterList`)
    pub custom_converter_list: Vec<String>,
    /// Automatic trim for sheet names and cell text. (Java `autoTrim`)
    pub auto_trim: Option<bool>,
    /// Whether Excel 1904 date windowing is enabled. (Java `use1904windowing`)
    pub use1904windowing: Option<bool>,
    /// Locale used for date/number formatting. (Java `locale`)
    pub locale: Option<String>,
    /// Whether scientific notation is used. (Java `useScientificFormat`)
    pub use_scientific_format: Option<bool>,
    /// Field-cache location for reflection metadata. (Java `filedCacheLocation`)
    pub filed_cache_location: Option<CacheLocation>,
}

impl BasicParameter {
    /// Creates an empty parameter bag. (Java default constructor)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the configured header rows. (Java `getHead()`)
    #[must_use]
    pub fn head(&self) -> Option<&[Vec<String>]> {
        self.head.as_deref()
    }

    /// Returns the model type name. (Java `getClazz()`)
    #[must_use]
    pub fn clazz(&self) -> Option<&str> {
        self.clazz.as_deref()
    }

    /// Returns custom converter registrations. (Java `getCustomConverterList()`)
    #[must_use]
    pub fn custom_converter_list(&self) -> &[String] {
        &self.custom_converter_list
    }
}
