//! Mirrors Java com.alibaba.excel.util.WriteHandlerUtils.
//!
//! Java dispatches every `WriteHandler` lifecycle callback
//! (`beforeWorkbookCreate`, `afterWorkbookCreate`, `beforeSheetCreate`,
//! ... `afterCellDispose`) through this helper class. The Rust port
//! replaces POI with `rust_xlsxwriter` and uses the
//! `WriteHandler` trait directly, so these helpers are no-op anchors
//! that preserve the 1:1 Java file mapping.

#![allow(dead_code)]

use std::any::Any;

type Ctx = Option<Box<dyn Any>>;

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createWorkbookWriteHandlerContext`.
#[must_use]
pub fn create_workbook_write_handler_context() -> Ctx {
    None
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeWorkbookCreate`.
pub fn before_workbook_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterWorkbookCreate`.
pub fn after_workbook_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterWorkbookDispose`.
pub fn after_workbook_dispose(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createSheetWriteHandlerContext`.
#[must_use]
pub fn create_sheet_write_handler_context() -> Ctx {
    None
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeSheetCreate`.
pub fn before_sheet_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterSheetCreate`.
pub fn after_sheet_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createCellWriteHandlerContext`.
#[must_use]
pub fn create_cell_write_handler_context() -> Ctx {
    None
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeCellCreate`.
pub fn before_cell_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterCellCreate`.
pub fn after_cell_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterCellDataConverted`.
pub fn after_cell_data_converted(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterCellDispose`.
pub fn after_cell_dispose(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#createRowWriteHandlerContext`.
#[must_use]
pub fn create_row_write_handler_context() -> Ctx {
    None
}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#beforeRowCreate`.
pub fn before_row_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.exexcel.util.WriteHandlerUtils#afterRowCreate`.
pub fn after_row_create(_ctx: &mut Ctx) {}

/// Mirrors `com.alibaba.excel.util.WriteHandlerUtils#afterRowDispose`.
pub fn after_row_dispose(_ctx: &mut Ctx) {}
