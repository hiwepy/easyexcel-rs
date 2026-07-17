//! Mirrors Java `com.alibaba.excel.enums.WriteTypeEnum`.
//!
//! `ADD` vs `FILL`. Used internally by `ExcelBuilderImpl` (Java) to switch
//! between `ExcelWriteAddExecutor` and `ExcelWriteFillExecutor`.

/// Write mode flag.
///
/// Rust port of Java `WriteTypeEnum`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WriteType {
    /// Append new rows. (Java `ADD`)
    Add,
    /// Fill template placeholders. (Java `FILL`)
    Fill,
}
