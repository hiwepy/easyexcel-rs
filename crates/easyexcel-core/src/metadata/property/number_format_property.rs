//! Mirrors Java `com.alibaba.excel.metadata.property.NumberFormatProperty`.

use bigdecimal::RoundingMode;

/// Java `java.math.RoundingMode` equivalent used by `@NumberFormat`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum NumberRoundingMode {
    /// Away from zero.
    Up,
    /// Toward zero.
    Down,
    /// Toward positive infinity.
    Ceiling,
    /// Toward negative infinity.
    Floor,
    /// Nearest value, ties away from zero.
    #[default]
    HalfUp,
    /// Nearest value, ties toward zero.
    HalfDown,
    /// Nearest value, ties to the even neighbour.
    HalfEven,
    /// Reject values that would require rounding.
    Unnecessary,
}

impl NumberRoundingMode {
    /// Returns the corresponding `bigdecimal` mode, or `None` for Java
    /// `UNNECESSARY`, which is enforced by comparing the rounded value.
    #[must_use]
    pub const fn bigdecimal(self) -> Option<RoundingMode> {
        match self {
            Self::Up => Some(RoundingMode::Up),
            Self::Down => Some(RoundingMode::Down),
            Self::Ceiling => Some(RoundingMode::Ceiling),
            Self::Floor => Some(RoundingMode::Floor),
            Self::HalfUp => Some(RoundingMode::HalfUp),
            Self::HalfDown => Some(RoundingMode::HalfDown),
            Self::HalfEven => Some(RoundingMode::HalfEven),
            Self::Unnecessary => None,
        }
    }
}

impl From<RoundingMode> for NumberRoundingMode {
    fn from(value: RoundingMode) -> Self {
        match value {
            RoundingMode::Up => Self::Up,
            RoundingMode::Down => Self::Down,
            RoundingMode::Ceiling => Self::Ceiling,
            RoundingMode::Floor => Self::Floor,
            RoundingMode::HalfUp => Self::HalfUp,
            RoundingMode::HalfDown => Self::HalfDown,
            RoundingMode::HalfEven => Self::HalfEven,
        }
    }
}

/// Number format metadata from `@NumberFormat`.
///
/// Rust port of Java `NumberFormatProperty`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormatProperty {
    /// Format pattern. (Java `format`)
    pub format: String,
    /// Rounding mode. (Java `roundingMode`)
    pub rounding_mode: NumberRoundingMode,
}

impl NumberFormatProperty {
    /// Creates a number format property. (Java constructor)
    #[must_use]
    pub fn new(format: impl Into<String>, rounding_mode: impl Into<NumberRoundingMode>) -> Self {
        Self {
            format: format.into(),
            rounding_mode: rounding_mode.into(),
        }
    }

    /// Builds from annotation values. (Java `build(NumberFormat)`)
    #[must_use]
    pub fn build(format: Option<&str>, rounding_mode: Option<NumberRoundingMode>) -> Option<Self> {
        format.map(|format| Self {
            format: format.to_owned(),
            rounding_mode: rounding_mode.unwrap_or_default(),
        })
    }

    /// Returns the format pattern. (Java `getFormat()`)
    #[must_use]
    pub fn format(&self) -> &str {
        &self.format
    }

    /// Returns the rounding mode. (Java `getRoundingMode()`)
    #[must_use]
    pub const fn rounding_mode(&self) -> NumberRoundingMode {
        self.rounding_mode
    }
}
