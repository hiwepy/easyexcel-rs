//! Mirrors Java `com.alibaba.excel.analysis.csv.CsvExcelReadExecutor`.

use crate::analysis::excel_read_executor::ExcelReadExecutor;
use crate::context::ReadSheet;

/// Mirrors Java `CsvExcelReadExecutor implements ExcelReadExecutor`.
///
/// The actual CSV parsing in Rust lives in `crate::read_csv`. This
/// struct exists for 1:1 Java package parity.
#[derive(Debug, Clone, Default)]
pub struct CsvExcelReadExecutor {
    /// Single logical sheet. (Java `sheetList`)
    sheet_list: Vec<ReadSheet>,
}

impl CsvExcelReadExecutor {
    /// Creates a new executor with the default CSV sheet.
    #[must_use]
    pub fn new() -> Self {
        Self {
            sheet_list: vec![ReadSheet::with_name(0, "Sheet1")],
        }
    }
}

impl ExcelReadExecutor for CsvExcelReadExecutor {
    /// Mirrors Java `sheetList()`.
    fn sheet_list(&self) -> &[ReadSheet] {
        &self.sheet_list
    }

    /// Mirrors Java `execute()`.
    fn execute(&mut self) {
        // Delegated to `read_csv` in the reader facade.
    }
}
