//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.sax.SharedStringsTableHandler`.
//!
//! Java's SAX handler parses `sharedStrings.xml` and populates the `ReadCache`.
//! Rust's `xlsx_rows::parse_shared_strings` drives this handler from `quick_xml`
//! events (same tag semantics: `si` / `t` / `rPh`).

use easyexcel_core::constant::excel_xml_constants::{
    SHAREDSTRINGS_RPH_TAG, SHAREDSTRINGS_SI_TAG, SHAREDSTRINGS_T_TAG,
};

/// SAX state machine for `xl/sharedStrings.xml`.
///
/// Corresponds to Java `SharedStringsTableHandler` fields
/// (`currentData`, `currentElementData`, `ignoreTagt`, `isTagt`).
#[derive(Debug, Default)]
pub struct SharedStringsTableHandler {
    /// Accumulated text for the current `<si>` item. (Java `currentData`)
    current_data: Option<String>,
    /// Text for the current `<t>` element. (Java `currentElementData`)
    current_element_data: Option<String>,
    /// Ignore `<t>` while inside phonetic runs. (Java `ignoreTagt`)
    ignore_tag_t: bool,
    /// True while inside a `<t>` that should collect characters. (Java `isTagt`)
    is_tag_t: bool,
}

impl SharedStringsTableHandler {
    /// Creates an empty handler. (Java constructor `SharedStringsTableHandler(ReadCache)`)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `SharedStringsTableHandler.startElement(...)`.
    ///
    /// `local_name` is the unqualified tag (`si` / `t` / `rPh`). Prefixed forms
    /// (`x:si`, `ns2:t`, …) are normalized by the caller via [`local_tag`].
    pub fn start_element(&mut self, local_name: &str) {
        match local_name {
            SHAREDSTRINGS_T_TAG => {
                self.current_element_data = None;
                self.is_tag_t = true;
            }
            SHAREDSTRINGS_SI_TAG => {
                self.current_data = None;
            }
            SHAREDSTRINGS_RPH_TAG => {
                self.ignore_tag_t = true;
            }
            _ => {}
        }
    }

    /// Java `SharedStringsTableHandler.endElement(...)`.
    ///
    /// Returns `Some(decoded)` when a complete `<si>` closes (including empty
    /// string for a missing `<t>`), matching `readCache.put(...)`.
    pub fn end_element(&mut self, local_name: &str) -> Option<String> {
        match local_name {
            SHAREDSTRINGS_T_TAG => {
                if let Some(element) = self.current_element_data.take() {
                    match &mut self.current_data {
                        Some(data) => data.push_str(&element),
                        None => self.current_data = Some(element),
                    }
                }
                self.is_tag_t = false;
                None
            }
            SHAREDSTRINGS_SI_TAG => {
                let raw = self.current_data.take().unwrap_or_default();
                Some(utf_decode(&raw))
            }
            SHAREDSTRINGS_RPH_TAG => {
                self.ignore_tag_t = false;
                None
            }
            _ => None,
        }
    }

    /// Java `SharedStringsTableHandler.characters(...)`.
    pub fn characters(&mut self, ch: &str) {
        if !self.is_tag_t || self.ignore_tag_t {
            return;
        }
        match &mut self.current_element_data {
            Some(buf) => buf.push_str(ch),
            None => self.current_element_data = Some(ch.to_owned()),
        }
    }
}

/// Strips XML namespace / `x:` / `ns2:` prefixes so switch arms match Java
/// `ExcelXmlConstants.SHAREDSTRINGS_*` local names.
#[must_use]
pub fn local_tag(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

/// Java `SharedStringsTableHandler.utfDecode(String)` — OOXML `_xHHHH_` escapes
/// (see OOXML §3.18.9 / POI `XSSFRichTextString`).
#[must_use]
pub fn utf_decode(value: &str) -> String {
    if !value.contains("_x") {
        return value.to_owned();
    }
    let bytes = value.as_bytes();
    let mut output = String::with_capacity(value.len());
    let mut index = 0;
    while index < bytes.len() {
        if index + 7 <= bytes.len()
            && bytes[index] == b'_'
            && bytes[index + 1] == b'x'
            && bytes[index + 6] == b'_'
            && let Ok(hex) = std::str::from_utf8(&bytes[index + 2..index + 6])
            && let Ok(code) = u16::from_str_radix(hex, 16)
            && let Some(character) = char::from_u32(u32::from(code))
        {
            output.push(character);
            index += 7;
        } else {
            let character = value[index..]
                .chars()
                .next()
                .expect("index is inside the UTF-8 string");
            output.push(character);
            index += character.len_utf8();
        }
    }
    output
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_simple_shared_string_item() {
        let mut handler = SharedStringsTableHandler::new();
        handler.start_element("si");
        handler.start_element("t");
        handler.characters("hello");
        assert_eq!(handler.end_element("t"), None);
        assert_eq!(handler.end_element("si").as_deref(), Some("hello"));
    }

    #[test]
    fn skips_phonetic_runs() {
        let mut handler = SharedStringsTableHandler::new();
        handler.start_element("si");
        handler.start_element("t");
        handler.characters("漢");
        assert_eq!(handler.end_element("t"), None);
        handler.start_element("rPh");
        handler.start_element("t");
        handler.characters("ignored");
        assert_eq!(handler.end_element("t"), None);
        assert_eq!(handler.end_element("rPh"), None);
        assert_eq!(handler.end_element("si").as_deref(), Some("漢"));
    }

    #[test]
    fn utf_decode_expands_excel_escapes() {
        assert_eq!(utf_decode("a_x000D_b"), "a\rb");
        assert_eq!(utf_decode("plain"), "plain");
    }
}
