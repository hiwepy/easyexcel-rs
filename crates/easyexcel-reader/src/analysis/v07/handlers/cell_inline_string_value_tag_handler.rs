//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.CellInlineStringValueTagHandler`.
//!
//! Java class is empty — it inherits `characters` from
//! `AbstractCellValueTagHandler` for the inline `<t>` tag.

use super::abstract_cell_value_tag_handler::AbstractCellValueTagHandler;
use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `CellInlineStringValueTagHandler` (inline string `<t>`).
#[derive(Debug, Default)]
pub struct CellInlineStringValueTagHandler {
    inner: AbstractCellValueTagHandler,
}

impl CellInlineStringValueTagHandler {
    /// Creates an idle inline-string handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Takes accumulated inline `<t>` text.
    pub fn take(&mut self) -> String {
        self.inner.take()
    }
}

impl XlsxTagHandler for CellInlineStringValueTagHandler {
    /// Java `CellInlineStringValueTagHandler` inherits empty `startElement` —
    /// multiple rich-text `<t>` runs append into the same `tempData`.
    fn start_element(&mut self, _name: &str, _attrs: &str) {}

    /// Java `AbstractCellValueTagHandler.characters`.
    fn characters(&mut self, ch: &str) {
        self.inner.append(ch);
    }
}
