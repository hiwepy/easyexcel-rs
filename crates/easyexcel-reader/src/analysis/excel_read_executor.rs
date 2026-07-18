//! Mirrors Java `com.alibaba.excel.analysis.ExcelReadExecutor` (interface).

/// Mirrors Java `ExcelReadExecutor`.
///
/// Java declares `sheetList()` and `execute()`. Rust's `read_xlsx` /
/// `read_xls` / `read_csv` functions cover the same contract.
pub trait ExcelReadExecutor {
    /// Executes the read. (Java `execute()`)
    fn execute(&mut self);
}
