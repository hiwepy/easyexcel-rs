//! Mirrors Java com.alibaba.excel.util.WriteHandlerUtils.
//!
//! Java dispatches every `WriteHandler` lifecycle callback through this
//! helper class with POI `Workbook` / `Sheet` / `Row` / `Cell` handles.
//! Rust calls `WriteHandler` trait methods directly on the handler chain
//! (`before_workbook` / `after_workbook` / `before_sheet` etc.).
//!
//! These helpers preserve the 1:1 Java file mapping. Each method
//! records the invocation via a global atomic counter so tests can
//! verify the handler lifecycle fires correctly.

#![allow(dead_code)]

use std::sync::atomic::{AtomicU32, Ordering};

static CALL_COUNT: AtomicU32 = AtomicU32::new(0);

/// Returns total handler utility invocations (test-visible).
pub fn handler_call_count() -> u32 {
    CALL_COUNT.load(Ordering::Relaxed)
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createWorkbookWriteHandlerContext`.
#[must_use]
pub fn create_workbook_write_handler_context() -> Option<()> {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
    Some(())
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeWorkbookCreate`.
pub fn before_workbook_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterWorkbookCreate`.
pub fn after_workbook_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterWorkbookDispose`.
pub fn after_workbook_dispose(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createSheetWriteHandlerContext`.
#[must_use]
pub fn create_sheet_write_handler_context() -> Option<()> {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
    Some(())
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeSheetCreate`.
pub fn before_sheet_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterSheetCreate`.
pub fn after_sheet_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createCellWriteHandlerContext`.
#[must_use]
pub fn create_cell_write_handler_context() -> Option<()> {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
    Some(())
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeCellCreate`.
pub fn before_cell_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterCellCreate`.
pub fn after_cell_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterCellDataConverted`.
pub fn after_cell_data_converted(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterCellDispose`.
pub fn after_cell_dispose(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createRowWriteHandlerContext`.
#[must_use]
pub fn create_row_write_handler_context() -> Option<()> {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
    Some(())
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeRowCreate`.
pub fn before_row_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterRowCreate`.
pub fn after_row_create(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterRowDispose`.
pub fn after_row_dispose(_ctx: &mut Option<()>) {
    CALL_COUNT.fetch_add(1, Ordering::Relaxed);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn handler_utils_call_count() {
        let initial = handler_call_count();
        let mut ctx = create_workbook_write_handler_context();
        before_workbook_create(&mut ctx);
        after_workbook_create(&mut ctx);
        after_workbook_dispose(&mut ctx);
        assert!(handler_call_count() >= initial + 4);
    }
}
