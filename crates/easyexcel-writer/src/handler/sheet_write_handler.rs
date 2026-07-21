//! Mirrors Java `com.alibaba.excel.write.handler.SheetWriteHandler`.

use std::sync::atomic::{AtomicU32, Ordering};
use easyexcel_core::WriteSheetContext;

static CALLS: AtomicU32 = AtomicU32::new(0);
pub fn sheet_handler_calls() -> u32 { CALLS.load(Ordering::Relaxed) }

pub trait SheetWriteHandler: easyexcel_core::WriteHandler {
    fn before_sheet_create(&mut self, _context: &WriteSheetContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
    fn after_sheet_create(&mut self, _context: &WriteSheetContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
}
