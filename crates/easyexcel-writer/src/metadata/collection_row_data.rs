//! Mirrors Java `com.alibaba.excel.write.metadata.CollectionRowData`.

/// Mirrors Java `CollectionRowData implements RowData`.
///
/// Java wraps a `Collection<?>` of raw values for a no-model row. The Rust
/// port is a tuple newtype that holds the same `Vec<CellValue>`. It implements
/// [`easyexcel_core::ExcelRow`], so it can enter both the public writer facade
/// and `ExcelWriteAddExecutor`.
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

impl easyexcel_core::ExcelRow for CollectionRowData {
    fn schema() -> &'static [easyexcel_core::ExcelColumn] {
        &[]
    }

    fn from_row(row: &easyexcel_core::RowData) -> easyexcel_core::Result<Self> {
        let actual_row = row
            .clone()
            .with_read_default_return(easyexcel_core::ReadDefaultReturn::ActualData);
        let dynamic =
            <easyexcel_core::DynamicRow as easyexcel_core::ExcelRow>::from_row(&actual_row)?;
        Ok(Self(
            <easyexcel_core::DynamicRow as easyexcel_core::ExcelRow>::to_row(&dynamic)?,
        ))
    }

    fn to_row(&self) -> easyexcel_core::Result<Vec<easyexcel_core::CellValue>> {
        Ok(self.0.clone())
    }
}
