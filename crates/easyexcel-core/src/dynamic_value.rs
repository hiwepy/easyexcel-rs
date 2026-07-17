//! Mirrors Java `com.alibaba.excel.enums.ReadDefaultReturnEnum`'s payload
//! values: `String` / `ActualData` (i.e. `CellValue`) / `ReadCellData`.

use crate::cell_value::CellValue;
use crate::read_cell_data::ReadCellData;

/// A type-safe value in a Java-compatible no-model row.
///
/// Java uses `Map<Integer, Object>` plus `ReadDefaultReturnEnum` to switch the
/// value kind; Rust enforces the kind via the `DynamicValue` enum so the
/// caller cannot accidentally mix scalars with rich cell metadata.
#[derive(Debug, Clone, PartialEq)]
pub enum DynamicValue {
    /// A missing column inserted to preserve physical indexes or head width.
    Null,
    /// Text returned by Java's default `STRING` mode.
    String(String),
    /// Scalar returned by Java's `ACTUAL_DATA` mode.
    ActualData(CellValue),
    /// Metadata returned by Java's `READ_CELL_DATA` mode.
    ReadCellData(ReadCellData),
}
