//! Mirrors Java `com.alibaba.excel.read.metadata.holder.ReadRowHolder`.

use std::collections::HashMap;

use easyexcel_core::CellValue;

/// Mirrors Java `ReadRowHolder implements Holder`.
#[derive(Debug, Clone)]
pub struct ReadRowHolder {
    /// Mirrors `ReadRowHolder.rowIndex`.
    pub row_index: i32,
    /// Mirrors `ReadRowHolder.cellMap`.
    pub cell_map: HashMap<usize, CellValue>,
}

impl ReadRowHolder {
    /// Mirrors Java constructor.
    pub fn new(row_index: i32, cell_map: HashMap<usize, CellValue>) -> Self {
        Self {
            row_index,
            cell_map,
        }
    }
}
