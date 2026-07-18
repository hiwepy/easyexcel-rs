//! Mirrors Java `com.alibaba.excel.write.metadata.MapRowData`.

use std::collections::BTreeMap;

use easyexcel_core::CellValue;

/// Mirrors Java `MapRowData implements RowData`.
///
/// Java wraps a `Map<Integer, ?>` keyed by physical column index. Rust uses
/// a `BTreeMap` for deterministic ordering.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct MapRowData(pub BTreeMap<usize, CellValue>);

impl MapRowData {
    /// Creates a `MapRowData` from a column-indexed map.
    #[must_use]
    pub fn new(values: BTreeMap<usize, CellValue>) -> Self {
        Self(values)
    }

    /// Returns the underlying map. (Java `getMap()` equivalent)
    #[must_use]
    pub fn values(&self) -> &BTreeMap<usize, CellValue> {
        &self.0
    }

    /// Returns whether the row is empty. (Java `RowData.isEmpty()`)
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}