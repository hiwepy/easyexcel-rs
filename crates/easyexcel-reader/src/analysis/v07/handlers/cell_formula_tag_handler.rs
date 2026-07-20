//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.CellFormulaTagHandler`.

use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `CellFormulaTagHandler`.
#[derive(Debug, Default)]
pub struct CellFormulaTagHandler {
    /// Accumulated formula text. (Java `XlsxReadSheetHolder.tempFormula`)
    pub temp_formula: String,
}

impl CellFormulaTagHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `CellFormulaTagHandler.startElement`.
    pub fn begin_formula(&mut self) {
        self.temp_formula.clear();
    }

    /// Java `CellFormulaTagHandler.endElement` — returns the formula string.
    pub fn finish_formula(&mut self) -> String {
        std::mem::take(&mut self.temp_formula)
    }
}

impl XlsxTagHandler for CellFormulaTagHandler {
    /// Java `CellFormulaTagHandler.startElement`.
    fn start_element(&mut self, name: &str, _attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local == "f" {
            self.begin_formula();
        }
    }

    /// Java `CellFormulaTagHandler.endElement`.
    fn end_element(&mut self, name: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local == "f" {
            let _ = self.finish_formula();
        }
    }

    /// Java `CellFormulaTagHandler.characters`.
    fn characters(&mut self, ch: &str) {
        self.temp_formula.push_str(ch);
    }
}
