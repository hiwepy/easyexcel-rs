//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.CountTagHandler`.
//!
//! Parses the worksheet `<dimension ref="A1:Z99">` attribute to recover an
//! approximate total row count (`PositionUtils.getRow(totalStr) + 1`).

use std::collections::HashMap;

use easyexcel_core::constant::excel_xml_constants::{ATTRIBUTE_REF, CELL_RANGE_SPLIT};
use easyexcel_core::{ExcelError, Result};

use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `CountTagHandler`.
#[derive(Debug, Default)]
pub struct CountTagHandler {
    /// Approximate total row number from the dimension ref. (Java sheet holder)
    pub approximate_total_row_number: Option<u32>,
}

impl CountTagHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `CountTagHandler.startElement`.
    pub fn parse_dimension_ref(ref_attr: &str) -> Result<u32> {
        let end = ref_attr
            .rsplit_once(CELL_RANGE_SPLIT)
            .map(|(_, end)| end)
            .unwrap_or(ref_attr);
        let row = row_from_cell_ref(end)?;
        Ok(row.saturating_add(1))
    }

    /// Applies dimension attributes onto this handler.
    pub fn start_dimension(&mut self, attrs: &HashMap<String, String>) -> Result<()> {
        let Some(reference) = attrs.get(ATTRIBUTE_REF) else {
            return Ok(());
        };
        self.approximate_total_row_number = Some(Self::parse_dimension_ref(reference)?);
        Ok(())
    }

    /// Converts Java `approximateTotalRowNumber` into a zero-based last row index
    /// for [`crate::xlsx_rows`] `last_explicit_row` (OOXML `r` is 1-based).
    ///
    /// `None` when the dimension tag was never seen or parsed as empty.
    #[must_use]
    pub fn last_explicit_row_index(&self) -> Option<u32> {
        self.approximate_total_row_number
            .filter(|&n| n > 0)
            .map(|n| n.saturating_sub(1))
    }
}

impl XlsxTagHandler for CountTagHandler {
    /// Java `CountTagHandler.startElement`.
    fn start_element(&mut self, name: &str, attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local != "dimension" {
            return;
        }
        let mut map = HashMap::new();
        for token in attrs.split_whitespace() {
            if let Some((key, value)) = token.split_once('=') {
                map.insert(key.to_owned(), value.to_owned());
            }
        }
        let _ = self.start_dimension(&map);
    }
}

/// Java `PositionUtils.getRow(String)` — zero-based row from an A1 token.
fn row_from_cell_ref(cell: &str) -> Result<u32> {
    let digits_start = cell
        .char_indices()
        .rev()
        .find(|(_, c)| !c.is_ascii_digit())
        .map(|(i, _)| i + 1)
        .unwrap_or(0);
    let row_part = &cell[digits_start..];
    if row_part.is_empty() {
        return Err(ExcelError::Format(format!(
            "dimension ref missing row: {cell}"
        )));
    }
    let one_based: u32 = row_part
        .parse()
        .map_err(|error| ExcelError::Format(format!("{error}")))?;
    Ok(one_based.saturating_sub(1))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_dimension_end_row() {
        assert_eq!(CountTagHandler::parse_dimension_ref("A1:C10").unwrap(), 10);
        assert_eq!(CountTagHandler::parse_dimension_ref("Z99").unwrap(), 99);
    }

    #[test]
    fn last_explicit_row_index_from_dimension() {
        let mut handler = CountTagHandler::new();
        assert_eq!(handler.last_explicit_row_index(), None);
        let mut attrs = HashMap::new();
        attrs.insert(ATTRIBUTE_REF.to_owned(), "A1:C10".to_owned());
        handler.start_dimension(&attrs).unwrap();
        assert_eq!(handler.approximate_total_row_number, Some(10));
        assert_eq!(handler.last_explicit_row_index(), Some(9));
    }
}
