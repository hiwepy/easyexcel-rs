//! Mirrors Java `com.alibaba.excel.enums.NumericCellTypeEnum`.
//!
//! POI-specific supplement; not surfaced publicly by Rust.

/// Supplements POI `CellType` so write paths can distinguish date from number.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumericCellType {
    /// Plain number.
    Number,
    /// Date encoded as a serial number.
    Date,
}
