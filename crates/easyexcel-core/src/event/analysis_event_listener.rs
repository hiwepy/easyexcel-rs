//! Mirrors Java `com.alibaba.excel.event.AnalysisEventListener`.

use crate::CellValue;

pub trait AnalysisEventListener<T>: crate::ReadListener<T> {
    fn invoke_head_map(
        &mut self,
        head_map: std::collections::HashMap<usize, String>,
        context: &crate::AnalysisContext,
    ) {
        let _ = (head_map, context);
    }
    fn do_after_all_analysed(&mut self, context: &crate::AnalysisContext) -> crate::Result<()> {
        let _ = context;
        Ok(())
    }
}

fn _import_marker(v: CellValue) {
    let _ = v;
}
