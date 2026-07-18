//! Mirrors Java `com.alibaba.excel.write.metadata.RowData` (interface).

use easyexcel_core::CellValue;

/// Mirrors Java `RowData` interface (one method: `getCellValue(int)`).
///
/// Java models each cell of a basic-type row through a common interface so
/// `ExcelWriteAddExecutor` can branch on `CollectionRowData`, `MapRowData`,
/// or JavaBean row uniformly. Rust achieves the same uniformity by
/// accepting `&[CellValue]` slices from any source, so this trait is a
/// 1:1 API marker without runtime polymorphism.
pub trait RowData {
    /// Returns the cell value at the given column index. (Java `getCellValue(int)`)
    fn get_cell_value(&self, column_index: usize) -> Option<&CellValue>;

    /// Returns whether the row carries any value. (Java `isEmpty()`)
    fn is_empty(&self) -> bool;
}

impl RowData for [CellValue] {
    fn get_cell_value(&self, column_index: usize) -> Option<&CellValue> {
        self.get(column_index)
    }

    fn is_empty(&self) -> bool {
        self.is_empty()
    }
}

impl RowData for Vec<CellValue> {
    fn get_cell_value(&self, column_index: usize) -> Option<&CellValue> {
        self.get(column_index)
    }

    fn is_empty(&self) -> bool {
        self.is_empty()
    }
}