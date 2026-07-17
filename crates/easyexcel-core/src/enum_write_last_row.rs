//! Mirrors Java `com.alibaba.excel.enums.WriteLastRowTypeEnum`.
//!
//! Tracks whether a worksheet has been initialized with template data or
//! remains empty.

/// State of the worksheet's last row.
///
/// Rust port of Java `WriteLastRowTypeEnum`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WriteLastRow {
    /// Excel created without a template and nothing has been written.
    CommonEmpty,
    /// Excel created from a template and nothing has been written.
    TemplateEmpty,
    /// At least one row has been written.
    HasData,
}
