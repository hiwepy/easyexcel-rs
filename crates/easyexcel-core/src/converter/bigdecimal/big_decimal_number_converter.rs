//! Mirrors Java `com.alibaba.excel.converters.bigdecimal.BigDecimalNumberConverter`.
//!
//! The actual conversion logic lives in
//! `easyexcel-core/src/from_into_impls.rs`. This struct exists
//! for 1:1 Java package parity.

/// Mirrors Java `BigDecimalNumberConverter`.
#[allow(dead_code)]
pub struct BigDecimalNumberConverter;
