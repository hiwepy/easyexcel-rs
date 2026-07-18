//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.sax.SharedStringsTableHandler`.
//!
//! Java's SAX handler parses `sharedStrings.xml` and populates the
//! `ReadCache`. In Rust, the same parsing is performed by
//! `xlsx_rows.rs::read_shared_strings` using `quick_xml`. This struct
//! exists for 1:1 Java package parity.
#[allow(dead_code)]
pub struct SharedStringsTableHandler;
