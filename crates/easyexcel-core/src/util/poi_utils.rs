//! Mirrors Java com.alibaba.excel.util.PoiUtils.

#![allow(dead_code)]

/// Mirrors `com.alibaba.excel.util.PoiUtils#customHeight`.
///
/// Java reflects on `XSSFRow` / `HSSFRow` to read the `customHeight`
/// attribute. The Rust writer uses `rust_xlsxwriter`, which exposes
/// this directly via `Worksheet::set_row_height`, so the helper is an
/// inert placeholder that defaults to `false` to preserve the 1:1 file
/// mapping.
#[must_use]
pub fn custom_height() -> bool {
    false
}
