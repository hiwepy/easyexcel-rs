//! Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyser` (interface).

use easyexcel_core::AnalysisContext;

/// Mirrors Java `ExcelAnalyser`.
///
/// Java declares four methods: `analysis`, `finish`, `excelExecutor`,
/// `analysisContext`. Rust's `read_xlsx` / `read_xls` / `read_csv`
/// functions cover the same contract functionally; this trait exists
/// for 1:1 Java package parity.
pub trait ExcelAnalyser {
    /// Parses the specified sheets. (Java `analysis(List<ReadSheet>, Boolean)`)
    fn analysis(&mut self);

    /// Completes the read, releasing caches and closing streams. (Java `finish()`)
    fn finish(&mut self);

    /// Returns the analysis context. (Java `analysisContext()`)
    fn analysis_context(&self) -> &AnalysisContext;
}
