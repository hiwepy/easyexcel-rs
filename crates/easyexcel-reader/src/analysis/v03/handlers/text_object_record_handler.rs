//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.TextObjectRecordHandler`.
//!
//! Java's handler processes one BIFF record type. In Rust, XLS parsing
//! is delegated to `calamine::Xls` which materializes `Range<Data>`.
//! The cell-by-cell dispatch happens in `reader/lib.rs::read_range`.
//! This struct exists for 1:1 Java package parity.

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `TextObjectRecordHandler`.
#[allow(dead_code)]
pub struct TextObjectRecordHandler;

impl XlsRecordHandler for TextObjectRecordHandler {}
