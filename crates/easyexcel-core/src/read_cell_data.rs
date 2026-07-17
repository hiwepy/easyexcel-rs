//! Mirrors Java `com.alibaba.excel.metadata.data.ReadCellData`.

use crate::cell_value::CellValue;
use crate::formula_data::FormulaData;

/// Java-compatible no-model cell metadata.
///
/// Mirrors Java `ReadCellData<T>`: `rowIndex`, `columnIndex`, `numberValue`,
/// `originalNumberValue`, `stringValue`, `booleanValue`, `data`, `type`,
/// `dataFormatData`, `formulaData`. The Rust port preserves the read-side
/// metadata that downstream consumers need.
#[derive(Debug, Clone, PartialEq)]
pub struct ReadCellData {
    row_index: u32,
    column_index: usize,
    raw_value: CellValue,
    data: CellValue,
    display_value: String,
    formula: Option<FormulaData>,
}

impl ReadCellData {
    /// Internal constructor mirroring Java's `ReadCellData(type, stringValue)`.
    /// Not part of the public API.
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn new(
        row_index: u32,
        column_index: usize,
        raw_value: CellValue,
        data: CellValue,
        display_value: String,
        formula: Option<FormulaData>,
    ) -> Self {
        Self {
            row_index,
            column_index,
            raw_value,
            data,
            display_value,
            formula,
        }
    }

    /// Returns the physical zero-based row index. (Java `getRowIndex()`)
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the physical zero-based column index. (Java `getColumnIndex()`)
    #[must_use]
    pub const fn column_index(&self) -> usize {
        self.column_index
    }

    /// Returns the original backend-neutral cell value. (Java `CellData.getData()`)
    #[must_use]
    pub const fn raw_value(&self) -> &CellValue {
        &self.raw_value
    }

    /// Returns the Java `ACTUAL_DATA`-equivalent value. (Java `getData()` for non-string)
    #[must_use]
    pub const fn data(&self) -> &CellValue {
        &self.data
    }

    /// Returns the Java-compatible formatted display text. (Java `getStringValue()`)
    #[must_use]
    pub fn display_value(&self) -> &str {
        &self.display_value
    }

    /// Returns formula metadata when the cell contains a formula. (Java `getFormulaData()`)
    #[must_use]
    pub const fn formula(&self) -> Option<&FormulaData> {
        self.formula.as_ref()
    }
}
