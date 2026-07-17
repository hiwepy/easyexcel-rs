//! Mirrors Java `com.alibaba.excel.enums.RowTypeEnum`.
//!
//! Used to distinguish data rows from empty rows during SAX streaming.

/// The types of row.
///
/// Rust port of Java `RowTypeEnum`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RowType {
    /// Data row.                  (Java `DATA`)
    Data,
    /// Empty row (only empty cells). (Java `EMPTY`)
    Empty,
}
