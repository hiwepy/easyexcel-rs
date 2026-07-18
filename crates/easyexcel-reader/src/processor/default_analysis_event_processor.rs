//! Mirrors Java `com.alibaba.excel.read.processor.DefaultAnalysisEventProcessor`.

use easyexcel_core::AnalysisContext;

use crate::processor::analysis_event_processor::AnalysisEventProcessor;

/// Mirrors Java `DefaultAnalysisEventProcessor implements AnalysisEventProcessor`.
#[derive(Debug, Clone, Default)]
pub struct DefaultAnalysisEventProcessor;

impl AnalysisEventProcessor for DefaultAnalysisEventProcessor {
    fn extra(&mut self, _analysis_context: &AnalysisContext) {}
    fn end_row(&mut self, _analysis_context: &AnalysisContext) {}
    fn end_sheet(&mut self, _analysis_context: &AnalysisContext) {}
}
