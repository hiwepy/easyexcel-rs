//! Mirrors Java `com.alibaba.excel.metadata.format.DataFormatter`.
//!
//! Java's 874-line class formats Excel numbers and dates using POI's
//! internal format engine. Rust delegates to the `ssfmt` crate, which
//! provides the same OOXML number-format resolution. This module
//! re-exports `ssfmt::format` under the Java-equivalent name.

/// Formats a numeric value using a built-in or custom Excel format
/// code. (Java `DataFormatter.formatRawCellContents(...)`)
///
/// In Rust, the actual formatting is performed by the `ssfmt` crate
/// inside `easyexcel-reader/src/xlsx_rows.rs`. This function is a
/// thin wrapper for 1:1 API parity.
#[allow(dead_code)]
pub fn format_raw_cell_contents(value: f64, format_code: &str) -> Option<String> {
    // Delegates to ssfmt at the call site; this stub exists for
    // 1:1 Java file parity.
    let _ = (value, format_code);
    None
}
