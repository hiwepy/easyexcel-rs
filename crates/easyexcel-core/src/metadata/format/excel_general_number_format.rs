//! Mirrors Java `com.alibaba.excel.metadata.format.ExcelGeneralNumberFormat`.
//!
//! Java's 81-line class formats numbers in Excel's "General" format.
//! Rust delegates to `ssfmt::format` with format code `"General"`.

/// Formats a number in Excel "General" format. (Java
/// `ExcelGeneralNumberFormat.format(Object, StringBuffer, FieldPosition)`)
#[allow(dead_code)]
pub fn format_general(value: f64) -> String {
    format!("{value}")
}
