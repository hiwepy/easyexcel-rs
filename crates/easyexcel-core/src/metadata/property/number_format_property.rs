//! Mirrors Java `com.alibaba.excel.metadata.property.NumberFormatProperty`.

use bigdecimal::RoundingMode;

/// Number format metadata from `@NumberFormat`.
///
/// Rust port of Java `NumberFormatProperty`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormatProperty {
    /// Format pattern. (Java `format`)
    pub format: String,
    /// Rounding mode. (Java `roundingMode`)
    pub rounding_mode: RoundingMode,
}

impl NumberFormatProperty {
    /// Creates a number format property. (Java constructor)
    #[must_use]
    pub fn new(format: impl Into<String>, rounding_mode: RoundingMode) -> Self {
        Self {
            format: format.into(),
            rounding_mode,
        }
    }

    /// Builds from annotation values. (Java `build(NumberFormat)`)
    #[must_use]
    pub fn build(format: Option<&str>, rounding_mode: Option<RoundingMode>) -> Option<Self> {
        format.map(|format| Self {
            format: format.to_owned(),
            rounding_mode: rounding_mode.unwrap_or(RoundingMode::HalfUp),
        })
    }

    /// Returns the format pattern. (Java `getFormat()`)
    #[must_use]
    pub fn format(&self) -> &str {
        &self.format
    }

    /// Returns the rounding mode. (Java `getRoundingMode()`)
    #[must_use]
    pub const fn rounding_mode(&self) -> RoundingMode {
        self.rounding_mode
    }
}
