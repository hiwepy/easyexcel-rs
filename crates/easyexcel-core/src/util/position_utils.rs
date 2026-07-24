//! Mirrors Java com.alibaba.excel.util.PositionUtils.

#![allow(dead_code)]

/// Mirrors `com.alibaba.excel.util.PositionUtils#getRowByRowTagt`.
///
/// Java reads the `r` attribute of an OOXML `<row>` element
/// (`rowTagt` was the internal name for that attribute). The value is
/// 1-based; this helper returns the same value as a `u32`.
#[must_use]
pub fn get_row_by_row_tagt(row_tag: &str) -> u32 {
    row_tag.parse::<u32>().unwrap_or(0)
}

/// Mirrors `com.alibaba.excel.util.PositionUtils#getRow`.
///
/// Java parses a 1-based row index from an `A1` style cell reference.
/// The result is returned 0-based to match the Rust internal indexing.
#[must_use]
pub fn get_row(cell_ref: &str) -> u32 {
    let digits: String = cell_ref
        .chars()
        .skip_while(|c| c.is_ascii_alphabetic())
        .collect();
    digits
        .parse::<u32>()
        .map(|n| n.saturating_sub(1))
        .unwrap_or(0)
}

/// Mirrors `com.alibaba.excel.util.PositionUtils#getCol`.
///
/// Java turns the leading letters of an `A1` reference into a 0-based
/// column index: `A` -> 0, `Z` -> 25, `AA` -> 26, ...
#[must_use]
pub fn get_col(cell_ref: &str) -> u32 {
    let mut col: u32 = 0;
    for c in cell_ref.chars().take_while(|c| c.is_ascii_alphabetic()) {
        col = col
            .saturating_mul(26)
            .saturating_add((c.to_ascii_uppercase() as u32).saturating_sub('A' as u32));
        col = col.saturating_add(1);
    }
    col.saturating_sub(1)
}
