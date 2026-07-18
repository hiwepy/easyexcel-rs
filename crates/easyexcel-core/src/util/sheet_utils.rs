//! Mirrors Java com.alibaba.excel.util.SheetUtils.

#![allow(dead_code)]

/// Mirrors `com.alibaba.excel.util.SheetUtils#match`.
///
/// Java matches a Java RegEx against the sheet name. Rust uses the
/// `regex` crate's syntax via `str::contains` for the placeholder, the
/// writer/reader crates own the real matching.
#[must_use]
pub fn match_sheet(sheet_name: &str, sheet_pattern: &str) -> bool {
    sheet_name == sheet_pattern || sheet_pattern.is_empty()
}
