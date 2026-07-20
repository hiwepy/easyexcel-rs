//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.CellTagHandler`.
//!
//! Attribute parsing and temp-buffer logic live here; `xlsx_rows::XlsxDisplayCellReader`
//! still owns the `quick_xml` event loop and calls into these helpers (只增不减).

use std::collections::HashMap;

use easyexcel_core::constant::excel_xml_constants::{ATTRIBUTE_R, ATTRIBUTE_S, ATTRIBUTE_T};
use easyexcel_core::{CellDataType, ExcelError, Result};

use super::xlsx_tag_handler::XlsxTagHandler;

/// Default style / format index when `c@s` is absent.
/// Java `CellTagHandler.DEFAULT_FORMAT_INDEX`.
const DEFAULT_FORMAT_INDEX: usize = 0;

/// Parsed `<c>` start attributes — used by both the handler and `xlsx_rows`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellStartAttrs {
    /// Zero-based `(row, column)` from `r` or the fallback cursor.
    pub position: (u32, usize),
    /// Zero-based style index from `s` (default 0).
    pub style_index: usize,
    /// Raw OOXML `t` attribute (`s` / `n` / `b` / …).
    pub cell_type: Option<String>,
    /// Logical type from Java `CellDataTypeEnum.buildFromCellType`.
    pub data_type: CellDataType,
}

/// Mirrors Java `CellTagHandler`.
///
/// Holds the per-cell temp buffer that Java stores on `XlsxReadSheetHolder`
/// (`tempCellData` / `tempData`).
#[derive(Debug, Default)]
pub struct CellTagHandler {
    /// Current column cursor after the last `startElement`. (Java sheet holder)
    pub column_index: Option<usize>,
    /// Style index from the last `c@s`.
    pub style_index: usize,
    /// OOXML type code from the last `c@t`.
    pub cell_type: Option<String>,
    /// Logical type from [`CellDataType::build_from_cell_type`].
    pub data_type: CellDataType,
    /// Accumulated character data for `<v>` / inline text. (Java `tempData`)
    pub temp_data: String,
}

impl CellTagHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `CellTagHandler.startElement(XlsxReadContext, String, Attributes)`.
    ///
    /// Parses `r` / `t` / `s` and resets `temp_data`.
    pub fn start_cell(
        &mut self,
        attrs: &HashMap<String, String>,
        fallback_row: u32,
        fallback_column: usize,
    ) -> Result<CellStartAttrs> {
        let parsed = Self::parse_start(attrs, fallback_row, fallback_column)?;
        self.column_index = Some(parsed.position.1);
        self.style_index = parsed.style_index;
        self.cell_type = parsed.cell_type.clone();
        self.data_type = parsed.data_type;
        self.temp_data.clear();
        Ok(parsed)
    }

    /// Pure attribute parse shared with `xlsx_rows::next_cell` (no self mutation required).
    ///
    /// Corresponds to the attribute-reading portion of Java `startElement`.
    pub fn parse_start(
        attrs: &HashMap<String, String>,
        fallback_row: u32,
        fallback_column: usize,
    ) -> Result<CellStartAttrs> {
        let position = match attrs.get(ATTRIBUTE_R) {
            Some(reference) => parse_cell_reference(reference)?,
            None => (fallback_row, fallback_column),
        };
        let style_index = match attrs.get(ATTRIBUTE_S) {
            Some(value) if !value.is_empty() => value
                .parse::<usize>()
                .map_err(|error| ExcelError::Format(error.to_string()))?,
            _ => DEFAULT_FORMAT_INDEX,
        };
        let cell_type = attrs.get(ATTRIBUTE_T).cloned();
        let data_type = CellDataType::build_from_cell_type(cell_type.as_deref()).ok_or_else(|| {
            ExcelError::Format(format!(
                "unsupported XLSX cell type: {}",
                cell_type.as_deref().unwrap_or_default()
            ))
        })?;
        Ok(CellStartAttrs {
            position,
            style_index,
            cell_type,
            data_type,
        })
    }

    /// Java `AbstractCellValueTagHandler.characters` path when this handler
    /// owns the temp buffer (also used when `<v>` text arrives).
    pub fn append_characters(&mut self, ch: &str) {
        self.temp_data.push_str(ch);
    }

    /// Clears per-cell state after `endElement`. (Java puts cell into `cellMap`)
    pub fn reset_temp(&mut self) {
        self.temp_data.clear();
        self.cell_type = None;
        self.data_type = CellDataType::Empty;
        self.style_index = DEFAULT_FORMAT_INDEX;
    }
}

impl XlsxTagHandler for CellTagHandler {
    /// Java `CellTagHandler.startElement` — `attrs` is `key=value` pairs separated by spaces.
    fn start_element(&mut self, name: &str, attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local != "c" {
            return;
        }
        let map = parse_attr_pairs(attrs);
        let _ = self.start_cell(&map, 0, self.column_index.unwrap_or(0));
    }

    /// Java `CellTagHandler.endElement` — clears temp buffers after the cell closes.
    fn end_element(&mut self, name: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local == "c" {
            self.reset_temp();
        }
    }

    /// Java path routes characters through value handlers; we also accept them here
    /// when the handler is used stand-alone.
    fn characters(&mut self, ch: &str) {
        self.append_characters(ch);
    }
}

/// Parses a space-separated `key=value` attribute bag used by [`XlsxTagHandler`].
fn parse_attr_pairs(attrs: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    for token in attrs.split_whitespace() {
        if let Some((key, value)) = token.split_once('=') {
            map.insert(key.to_owned(), value.to_owned());
        }
    }
    map
}

/// Minimal A1 parser for handler-local use (mirrors `xlsx_rows::parse_cell_reference`
/// without creating a circular module dependency).
fn parse_cell_reference(reference: &str) -> Result<(u32, usize)> {
    const MAX_ROW: u32 = 1_048_576;
    const MAX_COL: usize = 16_384;
    let reference = reference.strip_prefix('$').unwrap_or(reference);
    let column_end = reference
        .find(|character: char| !character.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (column, row) = reference.split_at(column_end);
    let row = row.strip_prefix('$').unwrap_or(row);
    if column.is_empty() || row.is_empty() || !row.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(ExcelError::Format(format!(
            "invalid cell reference: {reference}"
        )));
    }
    let mut one_based_column = 0_usize;
    for letter in column.bytes() {
        one_based_column = one_based_column
            .checked_mul(26)
            .and_then(|value| {
                value.checked_add(usize::from(letter.to_ascii_uppercase() - b'A' + 1))
            })
            .ok_or_else(|| ExcelError::Format(format!("invalid cell reference: {reference}")))?;
    }
    if !(1..=MAX_COL).contains(&one_based_column) {
        return Err(ExcelError::Format(format!(
            "column index exceeds XLSX limits: {reference}"
        )));
    }
    let one_based_row: u32 = row
        .parse()
        .map_err(|error| ExcelError::Format(format!("{error}")))?;
    if !(1..=MAX_ROW).contains(&one_based_row) {
        return Err(ExcelError::Format(format!(
            "row index exceeds XLSX limits: {reference}"
        )));
    }
    Ok((one_based_row - 1, one_based_column - 1))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_start_reads_r_t_s() {
        let mut attrs = HashMap::new();
        attrs.insert("r".into(), "B2".into());
        attrs.insert("t".into(), "s".into());
        attrs.insert("s".into(), "3".into());
        let parsed = CellTagHandler::parse_start(&attrs, 0, 0).unwrap();
        assert_eq!(parsed.position, (1, 1));
        assert_eq!(parsed.style_index, 3);
        assert_eq!(parsed.cell_type.as_deref(), Some("s"));
        assert_eq!(parsed.data_type, CellDataType::String);
    }
}
