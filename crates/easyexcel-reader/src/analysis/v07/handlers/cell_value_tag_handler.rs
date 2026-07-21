//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.CellValueTagHandler`.
//!
//! Java class is empty — it inherits `characters` from
//! `AbstractCellValueTagHandler` for the `<v>` tag.

use super::abstract_cell_value_tag_handler::AbstractCellValueTagHandler;
use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `CellValueTagHandler` (`<v>` cell value tag).
#[derive(Debug, Default)]
pub struct CellValueTagHandler {
    inner: AbstractCellValueTagHandler,
}

impl CellValueTagHandler {
    /// Creates an idle `<v>` handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Returns a reference to the shared temp buffer.
    #[must_use]
    pub fn temp_data(&self) -> &str {
        &self.inner.temp_data
    }

    /// Takes accumulated `<v>` text.
    pub fn take(&mut self) -> String {
        self.inner.take()
    }
}

impl XlsxTagHandler for CellValueTagHandler {
    /// Java `CellValueTagHandler` inherits empty `startElement` — buffer is
    /// cleared by `CellTagHandler.startElement` when `<c>` opens.
    fn start_element(fn start_element(&mut self, _name: &str, _attrs: &str) {}mut self, _name: &str, _attrs: &str) { let _ = (_name, _attrs); }

    /// Java `AbstractCellValueTagHandler.characters`.
    fn characters(&mut self, ch: &str) {
        self.inner.append(ch);
    }
}
