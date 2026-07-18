//! Mirrors Java com.alibaba.excel.util.BeanMapUtils.
//!
//! Java uses CGLIB's `BeanMap` to expose POJO fields as a `Map<String,
//! Object>` for fast reflection-driven reads/writes. The Rust port
//! generates the same mapping at compile time through the
//! `easyexcel_derive::ExcelRow` macro, so this helper is an inert
//! placeholder that preserves the 1:1 Java file mapping.

#![allow(dead_code)]

use std::any::Any;

/// Mirrors `com.alibaba.excel.util.BeanMapUtils#create`.
#[must_use]
pub fn create(_bean: &dyn Any) -> Option<Box<dyn Any>> {
    None
}
