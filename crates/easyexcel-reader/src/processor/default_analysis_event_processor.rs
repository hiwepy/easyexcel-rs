//! Mirrors Java `com.alibaba.excel.read.processor.DefaultAnalysisEventProcessor`.

use easyexcel_core::AnalysisContext;

pub trait AnalysisEventProcessor {
    fn extra(&mut self, _analysis_context: &AnalysisContext) {
        let _ = _analysis_context;
    }
    fn end_row(&mut self, _analysis_context: &AnalysisContext) {
        let _ = _analysis_context;
    }
    fn end_sheet(&mut self, _analysis_context: &AnalysisContext) {
        let _ = _analysis_context;
    }
}

pub struct DefaultAnalysisEventProcessor;
impl AnalysisEventProcessor for DefaultAnalysisEventProcessor {}
