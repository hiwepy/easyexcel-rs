//! Mirrors Java `com.alibaba.excel.analysis.csv.CsvExcelReadExecutor`.

use std::path::PathBuf;

use easyexcel_core::{ExcelError, ExcelRow, ReadListener, Result};

use crate::analysis::excel_read_executor::ExcelReadExecutor;
use crate::context::ReadSheet;
use crate::{ReadOptions, read_csv};

/// Mirrors Java `CsvExcelReadExecutor implements ExcelReadExecutor`.
///
/// The actual CSV parsing in Rust lives in `crate::read_csv`. This
/// struct exists for 1:1 Java package parity.
#[derive(Debug, Clone, Default)]
pub struct CsvExcelReadExecutor {
    /// Single logical sheet. (Java `sheetList`)
    sheet_list: Vec<ReadSheet>,
    /// CSV input path supplied by `ExcelAnalyserImpl`.
    path: Option<PathBuf>,
}

impl CsvExcelReadExecutor {
    /// Creates a new executor with the default CSV sheet.
    #[must_use]
    pub fn new() -> Self {
        Self {
            sheet_list: vec![ReadSheet::with_name(0, "Sheet1")],
            path: None,
        }
    }

    /// Creates an executor bound to a real CSV input.
    #[must_use]
    pub fn from_path(path: impl Into<PathBuf>) -> Self {
        Self {
            sheet_list: vec![ReadSheet::with_name(0, "Sheet1")],
            path: Some(path.into()),
        }
    }
}

impl ExcelReadExecutor for CsvExcelReadExecutor {
    /// Mirrors Java `sheetList()`.
    fn sheet_list(&self) -> &[ReadSheet] {
        &self.sheet_list
    }

    /// Mirrors Java `execute()` through the real CSV record parser.
    fn execute<T, L>(&mut self, options: &ReadOptions, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        let path = self.path.as_deref().ok_or_else(|| {
            ExcelError::Format("CsvExcelReadExecutor requires an input path".to_owned())
        })?;
        read_csv::<T, L>(path, options, listener)
    }
}

#[cfg(test)]
mod tests {
    use std::io::Write;

    use easyexcel_core::{AnalysisContext, DynamicRow};
    use tempfile::NamedTempFile;

    use super::*;

    #[derive(Default)]
    struct CollectingListener {
        rows: Vec<DynamicRow>,
    }

    impl ReadListener<DynamicRow> for CollectingListener {
        fn invoke(&mut self, data: DynamicRow, _context: &AnalysisContext) -> Result<()> {
            self.rows.push(data);
            Ok(())
        }
    }

    #[test]
    fn trait_execute_runs_the_real_csv_parser() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "value")?;
        writeln!(file, "csv-row")?;
        let mut options = ReadOptions::default();
        options.head_row_number = 1;
        let mut executor = CsvExcelReadExecutor::from_path(file.path());
        let mut listener = CollectingListener::default();

        ExcelReadExecutor::execute::<DynamicRow, _>(&mut executor, &options, &mut listener)?;

        assert_eq!(listener.rows.len(), 1);
        assert_eq!(executor.sheet_list()[0].sheet_name(), "Sheet1");
        Ok(())
    }

    #[test]
    fn unbound_executor_reports_a_real_configuration_error() {
        let mut executor = CsvExcelReadExecutor::new();
        let mut listener = CollectingListener::default();
        let error = ExcelReadExecutor::execute::<DynamicRow, _>(
            &mut executor,
            &ReadOptions::default(),
            &mut listener,
        )
        .expect_err("unbound CSV executor must fail");
        assert!(error.to_string().contains("requires an input path"));
    }
}
