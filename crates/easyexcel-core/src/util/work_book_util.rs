//! Mirrors Java com.alibaba.excel.util.WorkBookUtil.
//!
//! Java wraps Apache POI `Workbook` / `Sheet` / `Row` / `Cell`
//! construction with EasyExcel-specific defaults (sheet name sanitisation,
//! data format propagation). The Rust writer delegates the same concerns
//! to `rust_xlsxwriter`, so these helpers are inert placeholders that
//! preserve the 1:1 Java file mapping.

#![allow(dead_code)]

use std::any::Any;

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createWorkBook`.
#[must_use]
pub fn create_work_book() -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createSheet`.
#[must_use]
pub fn create_sheet(_sheet_name: &str) -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createRow`.
#[must_use]
pub fn create_row(_row_index: u32) -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#createCell`.
#[must_use]
pub fn create_cell(_column_index: u32) -> Option<Box<dyn Any>> {
    None
}

/// Mirrors `com.alibaba.excel.util.WorkBookUtil#fillDataFormat`.
pub fn fill_data_format(_format: &str) -> Result<(), crate::excel_error::ExcelError> {
    Ok(())
}
