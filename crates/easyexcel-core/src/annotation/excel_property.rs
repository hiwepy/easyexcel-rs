//! Mirrors Java `com.alibaba.excel.annotation.ExcelProperty`.
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(...)]` attributes
//! replaces Java runtime annotation processing. This module exists
//! for 1:1 Java file parity.

/// Marker type mirroring Java `@ExcelProperty`.
#[allow(dead_code)]
pub struct ExcelProperty;
