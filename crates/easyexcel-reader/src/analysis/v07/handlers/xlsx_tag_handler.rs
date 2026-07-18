//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.XlsxTagHandler`.
//!
//! Java interface with `support/startElement/endElement/characters`.
//! In Rust, the SAX event loop is a single `quick_xml::Reader::read_event_into`
//! match in `xlsx_rows.rs::XlsxDisplayCellReader::next_cell`. This trait
//! exists for 1:1 Java package parity.

/// Mirrors Java `XlsxTagHandler`.
pub trait XlsxTagHandler {
    /// Whether this handler supports the current context. (Java `support(XlsxReadContext)`)
    fn support(&self) -> bool { true }

    /// Called on the opening tag. (Java `startElement(XlsxReadContext, String, Attributes)`)
    fn start_element(&mut self, _name: &str, _attrs: &str) {}

    /// Called on the closing tag. (Java `endElement(XlsxReadContext, String)`)
    fn end_element(&mut self, _name: &str) {}

    /// Called on character data. (Java `characters(XlsxReadContext, char[], int, int)`)
    fn characters(&mut self, _ch: &str) {}
}
