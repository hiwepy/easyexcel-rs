//! Mirrors Java `com.alibaba.excel.converters.string.StringErrorConverter`.
//!
//! The actual conversion logic lives in
//! `easyexcel-core/src/from_into_impls.rs`. This struct exists
//! for 1:1 Java package parity.

/// Mirrors Java `StringErrorConverter`.
#[allow(dead_code)]
pub struct StringErrorConverter;
