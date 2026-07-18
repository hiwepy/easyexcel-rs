//! Mirrors Java `com.alibaba.excel.converters.bytearray.BoxingByteArrayImageConverter`.
//!
//! The actual conversion logic lives in
//! `easyexcel-core/src/from_into_impls.rs`. This struct exists
//! for 1:1 Java package parity.

/// Mirrors Java `BoxingByteArrayImageConverter`.
#[allow(dead_code)]
pub struct BoxingByteArrayImageConverter;
