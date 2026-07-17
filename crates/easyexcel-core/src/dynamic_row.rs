//! Mirrors Java `com.alibaba.excel.metadata.data.ReadCellData` collected into a
//! `Map<Integer, Object>`.

use std::collections::BTreeMap;

use crate::dynamic_value::DynamicValue;

/// A no-model row keyed by zero-based physical column index.
///
/// Java uses `Map<Integer, Object>` because the value type depends on
/// `ReadDefaultReturnEnum`. Rust enforces the variant via `DynamicValue` and
/// uses a `BTreeMap` for deterministic column order.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DynamicRow(pub(crate) BTreeMap<usize, DynamicValue>);

impl DynamicRow {
    /// Creates a dynamic row from indexed values.
    #[must_use]
    pub const fn new(values: BTreeMap<usize, DynamicValue>) -> Self {
        Self(values)
    }

    /// Returns all indexed values in physical column order.
    #[must_use]
    pub const fn values(&self) -> &BTreeMap<usize, DynamicValue> {
        &self.0
    }

    /// Returns a value by zero-based physical column index. (Java `Map.get(Integer)`)
    #[must_use]
    pub fn get(&self, column_index: usize) -> Option<&DynamicValue> {
        self.0.get(&column_index)
    }

    /// Consumes the row and returns its ordered values.
    #[must_use]
    pub fn into_values(self) -> BTreeMap<usize, DynamicValue> {
        self.0
    }
}
