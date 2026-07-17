//! Mirrors Java `com.alibaba.excel.enums.ReadDefaultReturnEnum`.
//!
//! `STRING` (default) / `ACTUAL_DATA` / `READ_CELL_DATA`.

/// Value mode used when reading rows without a declared Rust model.
///
/// Rust port of Java `ReadDefaultReturnEnum`. Mirrors the same three modes
/// while the `Default` impl reproduces Java's `STRING` default.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ReadDefaultReturn {
    /// Convert every present cell to the text a user sees in the workbook. (Java `STRING`, default)
    #[default]
    String,
    /// Preserve the backend-neutral scalar type of each cell. (Java `ACTUAL_DATA`)
    ActualData,
    /// Return the scalar together with its raw value, location, and formula. (Java `READ_CELL_DATA`)
    ReadCellData,
}
