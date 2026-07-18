//! Mirrors Java com.alibaba.excel.util.IntUtils.

#![allow(dead_code)]

/// Mirrors `com.google.common.primitives.Ints#saturatedCast`.
///
/// Clamps a wider integer (`i64`) into `i32` instead of panicking on
/// overflow: values outside `i32::MIN..=i32::MAX` are clipped to the
/// nearest bound.
#[must_use]
pub fn saturated_cast(value: i64) -> i32 {
    if value > i32::MAX as i64 {
        i32::MAX
    } else if value < i32::MIN as i64 {
        i32::MIN
    } else {
        value as i32
    }
}
