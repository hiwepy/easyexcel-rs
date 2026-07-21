//! Mirrors Java `com.alibaba.excel.event.AbstractIgnoreExceptionReadListener`.

use std::collections::HashMap;
use crate::analysis_context::AnalysisContext;
use crate::cell_extra::CellExtra;
use crate::read_listener::ReadListener;

pub trait AbstractIgnoreExceptionReadListener<T>: ReadListener<T> {
    fn on_exception_silent(fn on_exception_silent(&mut self, _error: &crate::excel_error::ExcelError, _context: &AnalysisContext) {}mut self, _error: &crate::excel_error::ExcelError, _context: &AnalysisContext) { let _ = (_error, _context); }
    fn extra_silent(&mut self, extra: &CellExtra, context: &AnalysisContext) {
        let _ = (extra, context);
    }
}

fn _import_marker(_: HashMap<usize, String>) {
    let _ = _;
}
