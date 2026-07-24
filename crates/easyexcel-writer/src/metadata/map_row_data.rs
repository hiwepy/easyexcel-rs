//! Mirrors Java `com.alibaba.excel.write.metadata.MapRowData`.

use std::collections::BTreeMap;

use easyexcel_core::CellValue;

/// Mirrors Java `MapRowData implements RowData`.
///
/// Java wraps a `Map<Integer, ?>` and its `RowData` adapter reports
/// `map.size()` then calls `map.get(0..size)`. Rust preserves that exact,
/// occasionally surprising contiguous-key contract. Use
/// [`easyexcel_core::DynamicRow`] when sparse physical-column semantics are
/// desired instead.
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

impl easyexcel_core::ExcelRow for MapRowData {
    fn schema() -> &'static [easyexcel_core::ExcelColumn] {
        &[]
    }

    fn from_row(row: &easyexcel_core::RowData) -> easyexcel_core::Result<Self> {
        let actual_row = row
            .clone()
            .with_read_default_return(easyexcel_core::ReadDefaultReturn::ActualData);
        let dynamic =
            <easyexcel_core::DynamicRow as easyexcel_core::ExcelRow>::from_row(&actual_row)?;
        let values = <easyexcel_core::DynamicRow as easyexcel_core::ExcelRow>::to_row(&dynamic)?
            .into_iter()
            .enumerate()
            .collect();
        Ok(Self(values))
    }

    fn to_row(&self) -> easyexcel_core::Result<Vec<CellValue>> {
        Ok((0..self.0.len())
            .map(|index| self.0.get(&index).cloned().unwrap_or(CellValue::Empty))
            .collect())
    }
}
