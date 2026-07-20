//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.AbstractCellValueTagHandler`.

use super::abstract_xlsx_tag_handler::AbstractXlsxTagHandler;
use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `AbstractCellValueTagHandler extends AbstractXlsxTagHandler`.
///
/// Java only overrides `characters` to append into `tempData`. Concrete
/// handlers (`CellValueTagHandler`, `CellInlineStringValueTagHandler`) inherit
/// that behaviour.
#[derive(Debug, Default)]
pub struct AbstractCellValueTagHandler {
    /// Character accumulator mirroring Java `XlsxReadSheetHolder.tempData`.
    pub temp_data: String,
}

impl AbstractCellValueTagHandler {
    /// Creates an idle accumulator.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `AbstractCellValueTagHandler.characters`.
    pub fn append(&mut self, ch: &str) {
        self.temp_data.push_str(ch);
    }

    /// Takes and clears the accumulated text.
    pub fn take(&mut self) -> String {
        std::mem::take(&mut self.temp_data)
    }
}

impl XlsxTagHandler for AbstractCellValueTagHandler {
    /// Java `AbstractCellValueTagHandler.characters`.
    fn characters(&mut self, ch: &str) {
        self.append(ch);
    }
}

/// Marker that concrete value handlers extend the abstract base (Java inheritance).
#[allow(dead_code)]
pub type AbstractCellValueBase = AbstractXlsxTagHandler;
