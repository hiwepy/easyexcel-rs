//! Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyser` (interface).

use easyexcel_core::{AnalysisContext, ExcelRow, ReadListener, Result};

use super::excel_read_executor::ExcelReadExecutorKind;

/// Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyser`.
///
/// Java declares four methods: `analysis`, `finish`, `excelExecutor`,
/// `analysisContext`. Rust's [`crate::read_xlsx`] / [`crate::read_xls`] /
/// [`crate::read_csv`] functions cover the same contract functionally;
/// [`super::ExcelAnalyserImpl`] is the hot-path dispatcher that selects among
/// them. This trait exists for 1:1 Java package parity.
pub trait ExcelAnalyser {
    /// Runs the selected executor with a typed listener.
    ///
    /// Java stores erased listeners in `ReadWorkbook`; Rust passes the typed
    /// listener explicitly and applies sheet/read-all selection through
    /// `ReadOptions` before this call.
    fn analysis<T, L>(&mut self, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>;

    /// Completes the read, releasing caches and closing streams. (Java `finish()`)
    fn finish(&mut self);

    /// Returns the selected format-specific executor. (Java `excelExecutor()`)
    fn excel_executor(&self) -> &ExcelReadExecutorKind;

    /// Returns the analysis context. (Java `analysisContext()`)
    fn analysis_context(&self) -> &AnalysisContext;
}
