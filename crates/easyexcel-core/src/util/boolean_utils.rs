//! Mirrors Java com.alibaba.excel.util.BooleanUtils.

#![allow(dead_code)]

/// Mirrors `org.apache.commons.lang3.BooleanUtils#valueOf`.
#[must_use]
pub fn value_of(value: &str) -> bool {
    matches!(
        value.trim().to_ascii_lowercase().as_str(),
        "true" | "t" | "yes" | "y" | "on" | "1"
    )
}

/// Mirrors `org.apache.commons.lang3.BooleanUtils#isTrue`.
#[must_use]
pub fn is_true(value: Option<bool>) -> bool {
        matches!(value, Some(true))
}

/// Mirrors `org.apache.commons.lang3.BooleanUtils#isNotTrue`.
#[must_use]
pub fn is_not_true(value: Option<bool>) -> bool {
    !is_true(value)
}

/// Mirrors `org.apache.commons.lang3.BooleanUtils#isFalse`.
#[must_use]
pub fn is_false(value: Option<bool>) -> bool {
    matches!(value, Some(false))
}

/// Mirrors `org.apache.commons.lang3.BooleanUtils#isNotFalse`.
#[must_use]
pub fn is_not_false(value: Option<bool>) -> bool {
    !is_false(value)
}
