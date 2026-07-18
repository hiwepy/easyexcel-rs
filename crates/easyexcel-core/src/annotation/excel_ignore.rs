//! Mirrors Java `com.alibaba.excel.annotation.ExcelIgnore`.
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(...)]` attributes
//! replaces Java runtime annotation processing. This module exists
//! for 1:1 Java file parity.

/// Marker type mirroring Java `@ExcelIgnore`.
#[allow(dead_code)]
pub struct ExcelIgnore;
