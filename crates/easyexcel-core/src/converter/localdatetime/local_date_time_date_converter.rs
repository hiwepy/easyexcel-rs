//! Mirrors Java `com.alibaba.excel.converters.localdatetime.LocalDateTimeDateConverter`.
//!
//! The actual conversion logic lives in
//! `easyexcel-core/src/from_into_impls.rs`. This struct exists
//! for 1:1 Java package parity.

/// Mirrors Java `LocalDateTimeDateConverter`.
#[allow(dead_code)]
pub struct LocalDateTimeDateConverter;
