//! Mirrors Java `com.alibaba.excel.annotation.format.NumberFormat`.
//!
//! In Rust, `#[excel(format = "...")]` replaces this annotation.

use crate::NumberRoundingMode;

/// Number formatting annotation metadata.
///
/// Rust derive syntax uses
/// `#[excel(number_format = "...", rounding_mode = "HALF_UP")]`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormat {
    value: String,
    rounding_mode: NumberRoundingMode,
}

impl NumberFormat {
    /// Creates Java-default annotation metadata.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns a copy with a decimal format pattern.
    #[must_use]
    pub fn value(mut self, value: impl Into<String>) -> Self {
        self.value = value.into();
        self
    }

    /// Returns a copy with a Java-compatible rounding mode.
    #[must_use]
    pub const fn rounding_mode(mut self, value: NumberRoundingMode) -> Self {
        self.rounding_mode = value;
        self
    }

    /// Returns the Java `value()` pattern.
    #[must_use]
    pub fn pattern(&self) -> &str {
        &self.value
    }

    /// Returns Java `roundingMode()`.
    #[must_use]
    pub const fn rounding(&self) -> NumberRoundingMode {
        self.rounding_mode
    }
}

impl Default for NumberFormat {
    fn default() -> Self {
        Self {
            value: String::new(),
            rounding_mode: NumberRoundingMode::HalfUp,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::NumberFormat;
    use crate::NumberRoundingMode;

    #[test]
    fn java_defaults_and_builder_values_are_preserved() {
        let defaults = NumberFormat::default();
        assert_eq!(defaults.pattern(), "");
        assert_eq!(defaults.rounding(), NumberRoundingMode::HalfUp);

        let configured = NumberFormat::new()
            .value("#.##%")
            .rounding_mode(NumberRoundingMode::Unnecessary);
        assert_eq!(configured.pattern(), "#.##%");
        assert_eq!(configured.rounding(), NumberRoundingMode::Unnecessary);
    }
}
