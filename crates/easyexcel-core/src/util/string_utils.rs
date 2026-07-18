//! Mirrors Java com.alibaba.excel.util.StringUtils.

#![allow(dead_code)]

/// Mirrors `org.apache.commons.lang3.StringUtils#isEmpty`.
#[must_use]
pub fn is_empty(cs: Option<&str>) -> bool {
    match cs {
        Some(s) => s.is_empty(),
        None => true,
    }
}

/// Mirrors `org.apache.commons.lang3.StringUtils#isBlank`.
#[must_use]
pub fn is_blank(cs: Option<&str>) -> bool {
    match cs {
        Some(s) => s.trim().is_empty(),
        None => true,
    }
}

/// Mirrors `org.apache.commons.lang3.StringUtils#isNotBlank`.
#[must_use]
pub fn is_not_blank(cs: Option<&str>) -> bool {
    !is_blank(cs)
}

/// Mirrors `org.apache.commons.lang3.StringUtils#equals`.
#[must_use]
pub fn equals(cs1: Option<&str>, cs2: Option<&str>) -> bool {
    cs1 == cs2
}

/// Mirrors `java.lang.String#regionMatches(boolean, int, String, int, int)`.
#[must_use]
pub fn region_matches(
    ignore_case: bool,
    this_str: &str,
    this_offset: usize,
    other: &str,
    other_offset: usize,
    len: usize,
) -> bool {
    let this_chars: Vec<char> = this_str.chars().collect();
    let other_chars: Vec<char> = other.chars().collect();
    if this_offset + len > this_chars.len() || other_offset + len > other_chars.len() {
        return false;
    }
    for i in 0..len {
        let a = this_chars[this_offset + i];
        let b = other_chars[other_offset + i];
        if a == b {
            continue;
        }
        if !ignore_case {
            return false;
        }
        if a.to_ascii_lowercase() != b.to_ascii_lowercase() {
            return false;
        }
    }
    true
}

/// Mirrors `org.apache.commons.lang3.StringUtils#isNumeric`.
#[must_use]
pub fn is_numeric(cs: Option<&str>) -> bool {
    let s = match cs {
        Some(s) if !s.is_empty() => s,
        _ => return false,
    };
    s.chars().all(|c| c.is_ascii_digit())
}
