//! Mirrors Java `com.alibaba.excel.analysis.ExcelReadExecutor` (interface).

use crate::context::ReadSheet;

/// Mirrors Java `ExcelReadExecutor`.
///
/// Java declares `sheetList()` and `execute()`. Rust's `read_xlsx` /
/// `read_xls` / `read_csv` functions cover the same contract.
pub trait ExcelReadExecutor {
    /// Returns discovered worksheets. (Java `sheetList()`)
    fn sheet_list(&self) -> &[ReadSheet];

    /// Executes the read. (Java `execute()`)
    fn execute(&mut self);
}
