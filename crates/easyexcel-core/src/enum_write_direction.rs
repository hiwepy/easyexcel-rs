//! Mirrors Java `com.alibaba.excel.enums.WriteDirectionEnum`.
//!
//! `VERTICAL` vs `HORIZONTAL` for template fills.
//!
//! Java uses this enum; the `easyexcel-template` crate uses
//! `easyexcel_template::FillDirection` which already provides the same two
//! variants. This enum is kept as a type alias to avoid diverging names.

/// Direction in which a template fill expands.
///
/// Rust port of Java `WriteDirectionEnum`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WriteDirection {
    /// Expand downward.
    Vertical,
    /// Expand rightward.
    Horizontal,
}
