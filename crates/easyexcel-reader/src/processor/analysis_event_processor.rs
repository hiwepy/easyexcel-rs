//! Mirrors Java `com.alibaba.excel.read.processor.AnalysisEventProcessor`.

use easyexcel_core::AnalysisContext;

/// Mirrors Java `AnalysisEventProcessor` (interface, 33 lines).
///
/// Java declares three methods: `extra`, `endRow`, `endSheet`. Rust
/// mirrors the same surface so a default implementation can be plugged
/// in alongside the concrete `DefaultAnalysisEventProcessor`.
pub trait AnalysisEventProcessor {
    /// Called when extra metadata is encountered. (Java `extra(AnalysisContext)`)
    fn extra(&mut self, analysis_context: &AnalysisContext);

    /// Called at the end of every row. (Java `endRow(AnalysisContext)`)
    fn end_row(&mut self, analysis_context: &AnalysisContext);

    /// Called at the end of every sheet. (Java `endSheet(AnalysisContext)`)
    fn end_sheet(&mut self, analysis_context: &AnalysisContext);
}
