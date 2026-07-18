//! Mirrors Java com.alibaba.excel.util.MapUtils.

#![allow(dead_code)]

use std::collections::{BTreeMap, HashMap};

/// Stand-in for Java's `LinkedHashMap`.
///
/// `std::collections` has no exact `LinkedHashMap` (it was removed from
/// the std prelude long ago). For an Excel writer the insertion order
/// of keys matters, so the Rust port returns a `Vec<(K, V)>`-backed
/// map-via-`BTreeMap` is *not* what we want; callers that need real
/// insertion ordering should depend on the `linked-hash-map` or
/// `indexmap` crate. The 1:1 Java mirror keeps the helper available.
pub struct LinkedHashMap<K, V> {
    inner: Vec<(K, V)>,
}

impl<K: PartialEq, V> LinkedHashMap<K, V> {
    /// Creates an empty map.
    #[must_use]
    pub fn new() -> Self {
        Self { inner: Vec::new() }
    }

    /// Creates an empty map with the given capacity.
    #[must_use]
    pub fn with_capacity(capacity: usize) -> Self {
        Self {
            inner: Vec::with_capacity(capacity),
        }
    }
}

impl<K: PartialEq, V> Default for LinkedHashMap<K, V> {
    fn default() -> Self {
        Self::new()
    }
}

/// Mirrors `org.apache.commons.collections4.MapUtils#newHashMap`.
#[must_use]
pub fn new_hash_map<K, V>() -> HashMap<K, V> {
    HashMap::new()
}

/// Mirrors `com.google.common.collect.Maps#newHashMapWithExpectedSize`.
#[must_use]
pub fn new_hash_map_with_expected_size<K, V>(expected_size: usize) -> HashMap<K, V> {
    HashMap::with_capacity(expected_size)
}

/// Mirrors `org.apache.commons.collections4.MapUtils#newTreeMap`.
#[must_use]
pub fn new_tree_map<K, V>() -> BTreeMap<K, V> {
    BTreeMap::new()
}

/// Mirrors `org.apache.commons.collections4.MapUtils#newLinkedHashMap`.
#[must_use]
pub fn new_linked_hash_map<K: PartialEq, V>() -> LinkedHashMap<K, V> {
    LinkedHashMap::new()
}

/// Mirrors `com.google.common.collect.Maps#newLinkedHashMapWithExpectedSize`.
#[must_use]
pub fn new_linked_hash_map_with_expected_size<K: PartialEq, V>(
    expected_size: usize,
) -> LinkedHashMap<K, V> {
    LinkedHashMap::with_capacity(expected_size)
}
