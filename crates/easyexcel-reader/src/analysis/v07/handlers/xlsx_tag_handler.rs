//! Mirrors Java SAX ContentHandler for XLSX tag dispatch.

pub trait XlsxTagHandler {
    fn start_element(&mut self, name: &str, attrs: &str) {
        let _ = (name, attrs);
    }
    fn end_element(&mut self, name: &str) {
        let _ = name;
    }
    fn characters(&mut self, ch: &str) {
        let _ = ch;
    }
}
