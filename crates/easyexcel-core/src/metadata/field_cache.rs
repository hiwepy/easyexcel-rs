//! Mirrors Java `com.alibaba.excel.metadata.FieldCache`.

use std::collections::BTreeMap;

use super::field_wrapper::FieldWrapper;

/// Cached, sorted model fields.
///
/// Rust port of Java `FieldCache`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FieldCache {
    /// Fields sorted by column order, excluding ignored fields. (Java `sortedFieldMap`)
    pub sorted_field_map: BTreeMap<i32, FieldWrapper>,
    /// Fields that explicitly use `@ExcelProperty.index`. (Java `indexFieldMap`)
    pub index_field_map: BTreeMap<i32, FieldWrapper>,
}

impl FieldCache {
    /// Creates a field cache. (Java all-args constructor)
    #[must_use]
    pub fn new(
        sorted_field_map: BTreeMap<i32, FieldWrapper>,
        index_field_map: BTreeMap<i32, FieldWrapper>,
    ) -> Self {
        Self {
            sorted_field_map,
            index_field_map,
        }
    }

    /// Returns the sorted field map. (Java `getSortedFieldMap()`)
    #[must_use]
    pub fn sorted_field_map(&self) -> &BTreeMap<i32, FieldWrapper> {
        &self.sorted_field_map
    }

    /// Returns the index field map. (Java `getIndexFieldMap()`)
    #[must_use]
    pub fn index_field_map(&self) -> &BTreeMap<i32, FieldWrapper> {
        &self.index_field_map
    }
}
