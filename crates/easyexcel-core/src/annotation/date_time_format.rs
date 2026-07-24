//! Mirrors Java `com.alibaba.excel.annotation.format.DateTimeFormat`.
//!
//! In Rust, `#[excel(format = "...")]` replaces this annotation.

use crate::BooleanEnum;

/// Date formatting annotation metadata.
///
/// Rust derive syntax uses
/// `#[excel(format = "...", use_1904_windowing = true)]`; this value type
/// preserves Java's reflective annotation surface for metadata consumers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTimeFormat {
    value: String,
    use_1904_windowing: BooleanEnum,
}

impl DateTimeFormat {
    /// Creates Java-default annotation metadata.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns a copy with a date format pattern.
    #[must_use]
    pub fn value(mut self, value: impl Into<String>) -> Self {
        self.value = value.into();
        self
    }

    /// Returns a copy with explicit 1900/1904 windowing metadata.
    #[must_use]
    pub const fn use_1904_windowing(mut self, value: BooleanEnum) -> Self {
        self.use_1904_windowing = value;
        self
    }

    /// Returns the Java `value()` pattern.
    #[must_use]
    pub fn pattern(&self) -> &str {
        &self.value
    }

    /// Returns Java `use1904windowing()`.
    #[must_use]
    pub const fn windowing(&self) -> BooleanEnum {
        self.use_1904_windowing
    }
}

impl Default for DateTimeFormat {
    fn default() -> Self {
        Self {
            value: String::new(),
            use_1904_windowing: BooleanEnum::Default,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::DateTimeFormat;
    use crate::BooleanEnum;

    #[test]
    fn java_defaults_and_builder_values_are_preserved() {
        let defaults = DateTimeFormat::default();
        assert_eq!(defaults.pattern(), "");
        assert_eq!(defaults.windowing(), BooleanEnum::Default);

        let configured = DateTimeFormat::new()
            .value("%Y-%m-%d")
            .use_1904_windowing(BooleanEnum::True);
        assert_eq!(configured.pattern(), "%Y-%m-%d");
        assert_eq!(configured.windowing(), BooleanEnum::True);
    }
}
