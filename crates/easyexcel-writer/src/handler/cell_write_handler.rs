//! Mirrors Java `com.alibaba.excel.write.handler.CellWriteHandler`.

use std::sync::atomic::{AtomicU32, Ordering};
use easyexcel_core::{WriteCellContext, WriteRowContext};

static CALLS: AtomicU32 = AtomicU32::new(0);
pub fn cell_handler_calls() -> u32 { CALLS.load(Ordering::Relaxed) }

pub trait CellWriteHandler: easyexcel_core::WriteHandler {
    fn before_cell_create(&mut self, _context: &mut WriteCellContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
    fn after_cell_data_converted(&mut self, _context: &WriteCellContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
    fn after_cell_dispose(&mut self, _context: &WriteCellContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
    fn after_row_dispose_marker(&self) -> Option<&WriteRowContext> { None }
}
