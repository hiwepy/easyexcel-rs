//! Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyserImpl`.

use easyexcel_core::AnalysisContext;
use easyexcel_core::ExcelError;

use crate::analysis::excel_analyser::ExcelAnalyser;

/// Mirrors Java `ExcelAnalyserImpl implements ExcelAnalyser`.
///
/// The Java side dispatches to `XlsxSaxAnalyser`, `XlsSaxAnalyser`, or
/// `CsvExcelReadExecutor` based on `ExcelTypeEnum`. Rust achieves the
/// same dispatch in the `read_xlsx` / `read_xls` / `read_csv` free
/// functions. This struct exists for 1:1 Java package parity.
#[derive(Default)]
pub struct ExcelAnalyserImpl {
    finished: bool,
}

impl ExcelAnalyserImpl {
    /// Creates a new analyser. (Java `ExcelAnalyserImpl(ReadWorkbook)`)
    #[must_use]
    pub const fn new() -> Self {
        Self { finished: false }
    }

    /// Mirrors Java `choiceExcelExecutor(ReadWorkbook)`. In Rust, the
    /// path extension determines which read function is called.
    #[allow(dead_code)]
    fn choice_excel_executor(&self) -> Result<(), ExcelError> {
        Ok(())
    }
}

impl ExcelAnalyser for ExcelAnalyserImpl {
    fn analysis(&mut self) {
        // Delegated to `read_xlsx` / `read_xls` / `read_csv` in the
        // reader facade.
    }

    fn finish(&mut self) {
        self.finished = true;
    }

    fn analysis_context(&self) -> &AnalysisContext {
        unreachable!("analysis context is held by the reader dispatcher")
    }
}
