//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.MergeCellTagHandler`.

use std::collections::HashMap;

use easyexcel_core::constant::excel_xml_constants::ATTRIBUTE_REF;
use easyexcel_core::{CellExtra, CellExtraType, ExcelError, Result};

use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `MergeCellTagHandler`.
#[derive(Debug, Default)]
pub struct MergeCellTagHandler {
    /// Whether merge extras are enabled. (Java `support` / `extraReadSet`)
    pub enabled: bool,
    /// Last parsed merge extra (Java `setCellExtra` + `extra(...)`).
    pub last_extra: Option<CellExtra>,
}

impl MergeCellTagHandler {
    /// Creates a handler; `enabled` mirrors Java `support(XlsxReadContext)`.
    #[must_use]
    pub fn new(enabled: bool) -> Self {
        Self {
            enabled,
            last_extra: None,
        }
    }

    /// Java `MergeCellTagHandler.startElement`.
    pub fn start_merge(&mut self, attrs: &HashMap<String, String>) -> Result<()> {
        if !self.enabled {
            return Ok(());
        }
        let Some(reference) = attrs.get(ATTRIBUTE_REF) else {
            return Ok(());
        };
        if reference.is_empty() {
            return Ok(());
        }
        self.last_extra = Some(cell_extra_from_ref(CellExtraType::Merge, None, reference)?);
        Ok(())
    }

    /// Same as [`Self::start_merge`], but missing / empty `ref` is an error
    /// (matches historical `xlsx_rows::required_attribute` behaviour).
    pub fn start_merge_required(&mut self, attrs: &HashMap<String, String>) -> Result<()> {
        if !self.enabled {
            return Ok(());
        }
        let reference = attrs
            .get(ATTRIBUTE_REF)
            .filter(|value| !value.is_empty())
            .ok_or_else(|| ExcelError::Format("merge cell ref is missing".to_owned()))?;
        self.last_extra = Some(cell_extra_from_ref(CellExtraType::Merge, None, reference)?);
        Ok(())
    }
}

impl XlsxTagHandler for MergeCellTagHandler {
    fn support(&self) -> bool {
        self.enabled
    }

    /// Java `MergeCellTagHandler.startElement`.
    fn start_element(&mut self, name: &str, attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local != "mergeCell" {
            return;
        }
        let mut map = HashMap::new();
        for token in attrs.split_whitespace() {
            if let Some((key, value)) = token.split_once('=') {
                map.insert(key.to_owned(), value.to_owned());
            }
        }
        let _ = self.start_merge(&map);
    }
}

/// Builds a [`CellExtra`] from an A1 / A1:B2 reference (Java `new CellExtra(type, text, ref)`).
///
/// Also enforces first≤last ordering used by `xlsx_rows::parse_cell_range`.
pub(crate) fn cell_extra_from_ref(
    extra_type: CellExtraType,
    text: Option<String>,
    reference: &str,
) -> Result<CellExtra> {
    let (first, last) = match reference.split_once(':') {
        Some((first, last)) => (first, last),
        None => (reference, reference),
    };
    let (first_row, first_column) = parse_a1(first)?;
    let (last_row, last_column) = parse_a1(last)?;
    if first_row > last_row || first_column > last_column {
        return Err(ExcelError::Format(format!(
            "invalid cell range ordering: {reference}"
        )));
    }
    Ok(CellExtra::new(
        extra_type,
        text,
        first_row,
        last_row,
        first_column,
        last_column,
    ))
}

fn parse_a1(reference: &str) -> Result<(u32, usize)> {
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
