//! Mirrors Java `com.alibaba.excel.analysis.csv.CsvExcelReadExecutor`.

use crate::analysis::excel_read_executor::ExcelReadExecutor;

/// Mirrors Java `CsvExcelReadExecutor implements ExcelReadExecutor`.
///
/// The actual CSV parsing in Rust lives in `crate::read_csv`. This
/// struct exists for 1:1 Java package parity.
#[derive(Default)]
pub struct CsvExcelReadExecutor;

impl CsvExcelReadExecutor {
    /// Creates a new executor.
    #[must_use]
    pub const fn new() -> Self {
        Self
    }
}

impl ExcelReadExecutor for CsvExcelReadExecutor {
    fn execute(&mut self) {
        // Delegated to `read_csv` in the reader facade.
    }
}
