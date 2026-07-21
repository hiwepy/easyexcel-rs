//! Mirrors Java `com.alibaba.excel.write.handler.WorkbookWriteHandler`.

use std::sync::atomic::{AtomicU32, Ordering};

use easyexcel_core::WriteWorkbookContext;

static CALLS: AtomicU32 = AtomicU32::new(0);

/// Returns total WorkbookWriteHandler lifecycle invocations (test-visible).
pub fn workbook_handler_calls() -> u32 { CALLS.load(Ordering::Relaxed) }

/// Mirrors Java `WorkbookWriteHandler extends WriteHandler`.
pub trait WorkbookWriteHandler: easyexcel_core::WriteHandler {
    /// Called before a workbook is created. (Java `beforeWorkbookCreate`)
    fn before_workbook_create(&mut self, _context: &WriteWorkbookContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }

    /// Called after a workbook is created. (Java `afterWorkbookCreate`)
    fn after_workbook_create(&mut self, _context: &WriteWorkbookContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }

    /// Called after the workbook has been disposed. (Java `afterWorkbookDispose`)
    fn after_workbook_dispose(&mut self, _context: &WriteWorkbookContext) {
        CALLS.fetch_add(1, Ordering::Relaxed);
    }
}
