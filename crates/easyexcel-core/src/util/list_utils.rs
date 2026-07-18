//! Mirrors Java com.alibaba.excel.util.ListUtils.

#![allow(dead_code)]

use std::vec::Vec;

/// Mirrors `org.apache.commons.collections4.ListUtils#newArrayList` /
/// the EasyExcel helper that wraps `new ArrayList<>()`.
#[must_use]
pub fn new_array_list<T>() -> Vec<T> {
    Vec::new()
}

/// Mirrors `com.alibaba.excel.util.ListUtils#newArrayListWithCapacity`.
#[must_use]
pub fn new_array_list_with_capacity<T>(capacity: usize) -> Vec<T> {
    Vec::with_capacity(capacity)
}

/// Mirrors `com.google.common.collect.Lists#newArrayListWithExpectedSize`.
#[must_use]
pub fn new_array_list_with_expected_size<T>(expected_size: usize) -> Vec<T> {
    // Guava's sizing: 1.5 * expected + 1, capped by isize::MAX.
    let cap = expected_size + (expected_size >> 1) + 1;
    Vec::with_capacity(cap)
}
