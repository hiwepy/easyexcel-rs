//! Mirrors Java `com.alibaba.excel.event.AnalysisEventListener`.

use crate::CellValue;

pub trait AnalysisEventListener<T>: crate::ReadListener<T> {
    fn invoke_head_map(fn invoke_head_map(&mut self, _head_map: std::collections::HashMap<usize, String>, _context: &crate::AnalysisContext) {}mut self, _head_map: std::collections::HashMap<usize, String>, _context: &crate::AnalysisContext) { let _ = (_head_map, _context); }
    fn do_after_all_analysed(&mut self, _context: &crate::AnalysisContext) -> crate::Result<()> { Ok(()) }
}

fn _import_marker(_: CellValue) {
    let _ = _;
}
