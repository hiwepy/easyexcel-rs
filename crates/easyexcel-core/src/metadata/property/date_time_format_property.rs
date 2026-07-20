//! Mirrors Java `com.alibaba.excel.metadata.property.DateTimeFormatProperty`.

/// Date-time format metadata from `@DateTimeFormat`.
///
/// Rust port of Java `DateTimeFormatProperty`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTimeFormatProperty {
    /// Format pattern. (Java `format`)
    pub format: String,
    /// Whether 1904 date windowing is enabled. (Java `use1904windowing`)
    pub use1904windowing: bool,
}

impl DateTimeFormatProperty {
    /// Creates a date-time format property. (Java constructor)
    #[must_use]
    pub fn new(format: impl Into<String>, use1904windowing: bool) -> Self {
        Self {
            format: format.into(),
            use1904windowing,
        }
    }

    /// Builds from annotation values. (Java `build(DateTimeFormat)`)
    #[must_use]
    pub fn build(format: Option<&str>, use1904windowing: Option<bool>) -> Option<Self> {
        format.map(|format| Self {
            format: format.to_owned(),
            use1904windowing: use1904windowing.unwrap_or(false),
        })
    }

    /// Returns the format pattern. (Java `getFormat()`)
    #[must_use]
    pub fn format(&self) -> &str {
        &self.format
    }

    /// Returns the 1904-windowing flag. (Java `getUse1904windowing()`)
    #[must_use]
    pub const fn use1904windowing(&self) -> bool {
        self.use1904windowing
    }
}
