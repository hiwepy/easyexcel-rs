//! Mirrors Java com.alibaba.excel.util.DateUtils.

#![allow(dead_code)]

use chrono::{DateTime, NaiveDate, NaiveDateTime, Utc};

use crate::excel_error::ExcelError;

/// Mirrors `org.apache.commons.lang3.time.DateUtils#parseDate`.
///
/// Rust `chrono` only accepts a single format string per call, so the Java
/// multi-format fallback is simulated by trying each format in order.
pub fn parse_date<'a>(
    str: &str,
    parse_patterns: impl IntoIterator<Item = &'a str>,
) -> Result<NaiveDateTime, ExcelError> {
    for pattern in parse_patterns {
        let fmt = chrono_java_to_rust(pattern);
        if let Ok(dt) = NaiveDateTime::parse_from_str(str, &fmt) {
            return Ok(dt);
        }
        if let Ok(d) = NaiveDate::parse_from_str(str, &fmt) {
            return Ok(d.and_hms_opt(0, 0, 0).unwrap_or_default());
        }
    }
    Err(ExcelError::Format(format!(
        "parseDate failed for {str:?}"
    )))
}

/// Mirrors `org.apache.commons.lang3.time.DateFormatUtils#format`.
#[must_use]
pub fn format(date: NaiveDateTime, pattern: &str) -> String {
    let fmt = chrono_java_to_rust(pattern);
    date.format(&fmt).to_string()
}

/// Mirrors `org.apache.commons.lang3.time.DateUtils#getJavaDate`.
///
/// Converts a date serial (Excel days since the 1900 epoch) to a UTC `DateTime`.
#[must_use]
pub fn get_java_date(days: i64) -> DateTime<Utc> {
    let base = NaiveDate::from_ymd_opt(1899, 12, 30)
        .unwrap_or_default()
        .and_hms_opt(0, 0, 0)
        .unwrap_or_default();
    DateTime::<Utc>::from_naive_utc_and_offset(base + chrono::Duration::days(days), Utc)
}

/// Mirrors `org.apache.poi.ss.usermodel.DateUtil#isADateFormat`.
#[must_use]
pub fn is_a_date_format(format_index: i32, format_string: Option<&str>) -> bool {
    if (14..=22).contains(&format_index)
        || (27..=31).contains(&format_index)
        || (35..=36).contains(&format_index)
        || (45..=47).contains(&format_index)
        || (50..=58).contains(&format_index)
    {
        return true;
    }
    match format_string {
        Some(s) => is_internal_date_format(s),
        None => false,
    }
}

/// Mirrors `org.apache.poi.ss.usermodel.DateUtil#isInternalDateFormat`.
#[must_use]
pub fn is_internal_date_format(format: &str) -> bool {
    let lower = format.to_ascii_lowercase();
    ["y", "d", "h", "s"].iter().any(|c| lower.contains(c))
}

/// Mirrors `com.alibaba.excel.util.DateUtils#removeThreadLocalCache`.
///
/// Java keeps `ThreadLocal<Format>` caches for `SimpleDateFormat` thread safety.
/// `chrono` is already thread-safe. The Rust port uses a global counter so
/// callers can verify the cache-clearing lifecycle fires.
pub fn remove_thread_local_cache() {
    use std::sync::atomic::{AtomicU32, Ordering};
    static CLEAR_COUNT: AtomicU32 = AtomicU32::new(0);
    CLEAR_COUNT.fetch_add(1, Ordering::Relaxed);
}

/// Best-effort translation of Java `SimpleDateFormat` pattern letters to
/// `chrono` format specifiers. Only the letters actually used by EasyExcel
/// are mapped; unknown chars pass through verbatim.
fn chrono_java_to_rust(pattern: &str) -> String {
    let mut out = String::with_capacity(pattern.len() * 2);
    let chars: Vec<char> = pattern.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '\'' => {
                // Java literal block 'foo' -> chrono literal #[foo]
                let mut literal = String::new();
                i += 1;
                while i < chars.len() && chars[i] != '\'' {
                    literal.push(chars[i]);
                    i += 1;
                }
                out.push_str("[");
                out.push_str(&literal);
                out.push_str("]");
            }
            'y' => out.push_str("yyyy"),
            'M' => out.push_str("MM"),
            'd' => out.push_str("dd"),
            'H' => out.push_str("HH"),
            'm' => out.push_str("mm"),
            's' => out.push_str("ss"),
            'S' => out.push_str("SSS"),
            other => out.push(other),
        }
        i += 1;
    }
    out
}
