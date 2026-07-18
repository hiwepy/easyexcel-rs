//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.AbstractXlsxTagHandler`.

use super::super::handlers::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `AbstractXlsxTagHandler implements XlsxTagHandler`.
///
/// Java provides default no-op implementations for all four methods.
/// Rust mirrors the same pattern.
#[allow(dead_code)]
pub struct AbstractXlsxTagHandler;

impl XlsxTagHandler for AbstractXlsxTagHandler {}
