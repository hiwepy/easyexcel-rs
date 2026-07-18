//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.sax.XlsxRowHandler`.
//!
//! Java's `DefaultHandler` subclass routes each XML tag to a
/// `XlsxTagHandler` from a lookup map. In Rust, the equivalent routing
/// is a `match` statement inside
/// `xlsx_rows.rs::XlsxDisplayCellReader::next_cell`. This struct exists
/// for 1:1 Java package parity.
#[allow(dead_code)]
pub struct XlsxRowHandler;
