//! Mirrors Java `com.alibaba.excel.write.metadata.CollectionRowData`.

/// Mirrors Java `CollectionRowData implements RowData`.
///
/// Java wraps a `Collection<?>` of raw values for a no-model row. The Rust
/// port is a tuple newtype that holds the same `Vec<CellValue>`; the
/// `ExcelWriteAddExecutor` consumes it as a slice.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CollectionRowData(pub Vec<easyexcel_core::CellValue>);

impl CollectionRowData {
    /// Creates a `CollectionRowData` mirroring Java's constructor.
    #[must_use]
    pub fn new(values: Vec<easyexcel_core::CellValue>) -> Self {
        Self(values)
    }

    /// Returns the underlying values. (Java `getCollection()` equivalent)
    #[must_use]
    pub fn values(&self) -> &[easyexcel_core::CellValue] {
        &self.0
    }

    /// Returns whether the row is empty. (Java `RowData.isEmpty()`)
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}