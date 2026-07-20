//! Mirrors Java `com.alibaba.excel.metadata.GlobalConfiguration`.

use crate::CacheLocation;

/// Global read/write configuration carried by holders.
///
/// Rust port of Java `GlobalConfiguration`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GlobalConfiguration {
    /// Automatic trim for sheet names and cell text. (Java `autoTrim`)
    pub auto_trim: bool,
    /// Whether Excel 1904 date windowing is enabled. (Java `use1904windowing`)
    pub use1904windowing: bool,
    /// Locale used for date/number formatting. (Java `locale`)
    pub locale: String,
    /// Whether scientific notation is used. (Java `useScientificFormat`)
    pub use_scientific_format: bool,
    /// Field-cache location for reflection metadata. (Java `filedCacheLocation`)
    pub filed_cache_location: CacheLocation,
}

impl Default for GlobalConfiguration {
    /// Mirrors Java default constructor values.
    fn default() -> Self {
        Self {
            auto_trim: true,
            use1904windowing: false,
            locale: "default".to_owned(),
            use_scientific_format: false,
            filed_cache_location: CacheLocation::ThreadLocal,
        }
    }
}

impl GlobalConfiguration {
    /// Creates a global configuration with Java default values. (Java constructor)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns the auto-trim flag. (Java `getAutoTrim()`)
    #[must_use]
    pub const fn auto_trim(&self) -> bool {
        self.auto_trim
    }

    /// Returns the 1904-windowing flag. (Java `getUse1904windowing()`)
    #[must_use]
    pub const fn use1904windowing(&self) -> bool {
        self.use1904windowing
    }

    /// Returns the locale name. (Java `getLocale()`)
    #[must_use]
    pub fn locale(&self) -> &str {
        &self.locale
    }

    /// Returns the scientific-format flag. (Java `getUseScientificFormat()`)
    #[must_use]
    pub const fn use_scientific_format(&self) -> bool {
        self.use_scientific_format
    }

    /// Returns the field-cache location. (Java `getFiledCacheLocation()`)
    #[must_use]
    pub const fn filed_cache_location(&self) -> CacheLocation {
        self.filed_cache_location
    }
}
