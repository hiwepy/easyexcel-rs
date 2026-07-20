//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.AbstractXlsxTagHandler`.

use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `AbstractXlsxTagHandler implements XlsxTagHandler`.
///
/// Java provides default no-op implementations for all four methods
/// (`support` / `startElement` / `endElement` / `characters`). Rust mirrors
/// the same pattern via trait defaults on [`XlsxTagHandler`].
#[derive(Debug, Default)]
pub struct AbstractXlsxTagHandler;

impl AbstractXlsxTagHandler {
    /// Creates the abstract base (rarely constructed on its own).
    #[must_use]
    pub fn new() -> Self {
        Self
    }
}

impl XlsxTagHandler for AbstractXlsxTagHandler {}
