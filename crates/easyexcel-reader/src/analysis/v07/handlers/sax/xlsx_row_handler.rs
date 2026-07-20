//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.sax.XlsxRowHandler`.
//!
//! Java routes each worksheet tag through a static `XLSX_CELL_HANDLER_MAP`.
//! This Rust port keeps the same map and forwards SAX-style callbacks; the
//! production reader still uses `xlsx_rows::XlsxDisplayCellReader` as the
//! primary event loop and may call individual handlers from that path.

use std::collections::HashMap;

use easyexcel_core::constant::excel_xml_constants::{
    CELL_FORMULA_TAG, CELL_INLINE_STRING_VALUE_TAG, CELL_TAG, CELL_VALUE_TAG, DIMENSION_TAG,
    HYPERLINK_TAG, MERGE_CELL_TAG, ROW_TAG,
};

use crate::analysis::v07::handlers::cell_formula_tag_handler::CellFormulaTagHandler;
use crate::analysis::v07::handlers::cell_inline_string_value_tag_handler::CellInlineStringValueTagHandler;
use crate::analysis::v07::handlers::cell_tag_handler::CellTagHandler;
use crate::analysis::v07::handlers::cell_value_tag_handler::CellValueTagHandler;
use crate::analysis::v07::handlers::count_tag_handler::CountTagHandler;
use crate::analysis::v07::handlers::hyperlink_tag_handler::HyperlinkTagHandler;
use crate::analysis::v07::handlers::merge_cell_tag_handler::MergeCellTagHandler;
use crate::analysis::v07::handlers::row_tag_handler::RowTagHandler;
use crate::analysis::v07::handlers::xlsx_tag_handler::XlsxTagHandler;

/// Tag → handler routing table, mirroring Java `XlsxRowHandler.XLSX_CELL_HANDLER_MAP`.
pub enum RoutedHandler {
    /// `<c>`
    Cell(CellTagHandler),
    /// `<row>`
    Row(RowTagHandler),
    /// `<v>`
    CellValue(CellValueTagHandler),
    /// inline `<t>`
    InlineString(CellInlineStringValueTagHandler),
    /// `<f>`
    Formula(CellFormulaTagHandler),
    /// `<dimension>`
    Count(CountTagHandler),
    /// `<mergeCell>`
    Merge(MergeCellTagHandler),
    /// `<hyperlink>`
    Hyperlink(HyperlinkTagHandler),
}

impl RoutedHandler {
    fn as_mut(&mut self) -> &mut dyn XlsxTagHandler {
        match self {
            Self::Cell(h) => h,
            Self::Row(h) => h,
            Self::CellValue(h) => h,
            Self::InlineString(h) => h,
            Self::Formula(h) => h,
            Self::Count(h) => h,
            Self::Merge(h) => h,
            Self::Hyperlink(h) => h,
        }
    }
}

/// Mirrors Java `XlsxRowHandler extends DefaultHandler`.
pub struct XlsxRowHandler {
    /// Active handlers keyed by local tag name.
    handlers: HashMap<&'static str, RoutedHandler>,
    /// Open-tag stack. (Java `XlsxReadSheetHolder.tagDeque`)
    tag_stack: Vec<String>,
}

impl XlsxRowHandler {
    /// Java `XlsxRowHandler(XlsxReadContext)` static map initialisation.
    #[must_use]
    pub fn new(read_merge: bool, read_hyperlink: bool) -> Self {
        let mut handlers = HashMap::new();
        handlers.insert(CELL_TAG, RoutedHandler::Cell(CellTagHandler::new()));
        handlers.insert(ROW_TAG, RoutedHandler::Row(RowTagHandler::new()));
        handlers.insert(CELL_VALUE_TAG, RoutedHandler::CellValue(CellValueTagHandler::new()));
        handlers.insert(
            CELL_INLINE_STRING_VALUE_TAG,
            RoutedHandler::InlineString(CellInlineStringValueTagHandler::new()),
        );
        handlers.insert(CELL_FORMULA_TAG, RoutedHandler::Formula(CellFormulaTagHandler::new()));
        handlers.insert(DIMENSION_TAG, RoutedHandler::Count(CountTagHandler::new()));
        handlers.insert(
            MERGE_CELL_TAG,
            RoutedHandler::Merge(MergeCellTagHandler::new(read_merge)),
        );
        handlers.insert(
            HYPERLINK_TAG,
            RoutedHandler::Hyperlink(HyperlinkTagHandler::new(read_hyperlink)),
        );
        Self {
            handlers,
            tag_stack: Vec::new(),
        }
    }

    /// Java `XlsxRowHandler.startElement`.
    pub fn start_element(&mut self, name: &str, attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        let Some(handler) = self.handlers.get_mut(local) else {
            return;
        };
        if !handler.as_mut().support() {
            return;
        }
        self.tag_stack.push(local.to_owned());
        handler.as_mut().start_element(name, attrs);
    }

    /// Java `XlsxRowHandler.characters`.
    pub fn characters(&mut self, ch: &str) {
        let Some(current) = self.tag_stack.last() else {
            return;
        };
        let key = current.clone();
        let Some(handler) = self.handlers.get_mut(key.as_str()) else {
            return;
        };
        if !handler.as_mut().support() {
            return;
        }
        handler.as_mut().characters(ch);
    }

    /// Java `XlsxRowHandler.endElement`.
    pub fn end_element(&mut self, name: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        let Some(handler) = self.handlers.get_mut(local) else {
            return;
        };
        if !handler.as_mut().support() {
            return;
        }
        handler.as_mut().end_element(name);
        let _ = self.tag_stack.pop();
    }
}
