//! Mirrors Java `com.alibaba.excel.event.AbstractIgnoreExceptionReadListener`.

use crate::analysis_context::AnalysisContext;
use crate::cell_extra::CellExtra;
use crate::read_listener::ReadListener;
use std::collections::HashMap;

pub trait AbstractIgnoreExceptionReadListener<T>: ReadListener<T> {
    fn on_exception_silent(
        &mut self,
        error: &crate::excel_error::ExcelError,
        context: &AnalysisContext,
    ) {
        let _ = (error, context);
    }
    fn extra_silent(&mut self, extra: &CellExtra, context: &AnalysisContext) {
        let _ = (extra, context);
    }
}

fn _import_marker(m: HashMap<usize, String>) {
    let _ = m;
}
