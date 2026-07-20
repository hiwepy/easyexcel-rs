//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.RowTagHandler`.
//!
//! Row-index resolution is shared with `xlsx_rows::XlsxDisplayCellReader::next_cell`
//! via [`RowTagHandler::resolve_row_index`]. Empty-row synthesis from Java
//! `startElement` (emitting `RowTypeEnum.EMPTY` for gaps) remains the
//! responsibility of higher-level read dispatchers.

use std::collections::HashMap;

use easyexcel_core::constant::excel_xml_constants::ATTRIBUTE_R;
use easyexcel_core::{ExcelError, Result};

use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `RowTagHandler`.
#[derive(Debug, Default)]
pub struct RowTagHandler {
    /// Zero-based current row index. (Java `XlsxReadSheetHolder.rowIndex`)
    pub row_index: Option<u32>,
    /// Whether the open row accumulated any non-empty cells. (Java `RowTypeEnum`)
    pub has_data: bool,
}

impl RowTagHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `RowTagHandler.startElement` — resolve `r` via
    /// `PositionUtils.getRowByRowTagt(rowTagt, before)`.
    ///
    /// Returns the zero-based row index for the opened `<row>`.
    pub fn resolve_row_index(row_attr: Option<&str>, before: u32) -> Result<u32> {
        match row_attr {
            Some(value) if !value.is_empty() => {
                let one_based: u32 = value
                    .parse()
                    .map_err(|error| ExcelError::Format(format!("{error}")))?;
                if !(1..=1_048_576).contains(&one_based) {
                    return Err(ExcelError::Format(format!(
                        "row index exceeds XLSX limits: {value}"
                    )));
                }
                Ok(one_based - 1)
            }
            _ => Ok(before),
        }
    }

    /// Java `RowTagHandler.startElement` body (without empty-row gap fill).
    pub fn start_row(&mut self, attrs: &HashMap<String, String>) -> Result<u32> {
        let before = self.row_index.unwrap_or(0);
        let row = Self::resolve_row_index(attrs.get(ATTRIBUTE_R).map(String::as_str), before)?;
        self.row_index = Some(row);
        self.has_data = false;
        Ok(row)
    }

    /// Java `RowTagHandler.endElement` — advances the cursor and reports whether
    /// the row looked like `DATA` vs `EMPTY`.
    pub fn end_row(&mut self) -> (u32, bool) {
        let row = self.row_index.unwrap_or(0);
        let has_data = self.has_data;
        self.row_index = Some(row.saturating_add(1));
        self.has_data = false;
        (row, has_data)
    }

    /// Marks that at least one non-empty cell was seen in the open row.
    pub fn mark_data(&mut self) {
        self.has_data = true;
    }
}

impl XlsxTagHandler for RowTagHandler {
    /// Java `RowTagHandler.startElement`.
    fn start_element(&mut self, name: &str, attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local != "row" {
            return;
        }
        let mut map = HashMap::new();
        for token in attrs.split_whitespace() {
            if let Some((key, value)) = token.split_once('=') {
                map.insert(key.to_owned(), value.to_owned());
            }
        }
        let _ = self.start_row(&map);
    }

    /// Java `RowTagHandler.endElement`.
    fn end_element(&mut self, name: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local == "row" {
            let _ = self.end_row();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_row_index_from_one_based_attr() {
        assert_eq!(RowTagHandler::resolve_row_index(Some("3"), 0).unwrap(), 2);
        assert_eq!(RowTagHandler::resolve_row_index(None, 4).unwrap(), 4);
    }
}
