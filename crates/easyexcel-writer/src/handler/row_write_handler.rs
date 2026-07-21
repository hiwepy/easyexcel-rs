//! Mirrors Java `com.alibaba.excel.write.handler.RowWriteHandler`.

use std::sync::atomic::{AtomicU32, Ordering};
use easyexcel_core::WriteRowContext;

static CALLS: AtomicU32 = AtomicU32::new(0);
pub fn row_handler_calls() -> u32 { CALLS.load(Ordering::Relaxed) }

pub trait RowWriteHandler: easyexcel_core::WriteHandler {
    fn before_row_create(&mut self, _context: &WriteRowContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
    fn after_row_create(&mut self, _context: &WriteRowContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
    fn after_row_dispose(&mut self, _context: &WriteRowContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
}
