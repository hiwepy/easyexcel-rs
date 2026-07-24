//! Mirrors Java `com.alibaba.excel.annotation.ExcelIgnore`.
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(...)]` attributes
//! replaces Java runtime annotation processing. This module exists
//! for 1:1 Java file parity.

/// Marker type mirroring Java `@ExcelIgnore`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ExcelIgnore;

impl ExcelIgnore {
    /// Creates the field-level ignore marker.
    #[must_use]
    pub const fn new() -> Self {
        Self
    }
}
