//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.RowTagHandler`.
//!
//! Java's handler processes one XML tag type inside the SAX event loop.
//! In Rust, the equivalent logic is inlined into the `quick_xml` event
//! match arms in `xlsx_rows.rs::XlsxDisplayCellReader`. This struct
//! exists for 1:1 Java package parity.

use super::super::handlers::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `RowTagHandler`.
#[allow(dead_code)]
pub struct RowTagHandler;

impl XlsxTagHandler for RowTagHandler {}
