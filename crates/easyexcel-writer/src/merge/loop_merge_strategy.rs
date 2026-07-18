//! Mirrors Java `com.alibaba.excel.write.merge.LoopMergeStrategy`.

use easyexcel_core::{CellExtra, WriteCellContext, WriteHandler};

use crate::merge::abstract_merge_strategy::AbstractMergeStrategy;

/// Mirrors Java `LoopMergeStrategy` (3 constructors + `afterRowDispose`).
pub struct LoopMergeStrategy {
    each_rows: u32,
    column_extend: u16,
    column_index: u16,
}

impl LoopMergeStrategy {
    /// Creates a `LoopMergeStrategy` with the given dimensions. (Java
    /// `LoopMergeStrategy(int eachRow, int columnExtend, int columnIndex)`)
    #[must_use]
    pub const fn new(each_rows: u32, column_extend: u16, column_index: u16) -> Self {
        Self {
            each_rows,
            column_extend,
            column_index,
        }
    }

    /// Returns the per-group row count. (Java `getEachRow()`)
    #[must_use]
    pub const fn each_rows(&self) -> u32 {
        self.each_rows
    }

    /// Returns the per-group column count. (Java `getColumnExtend()`)
    #[must_use]
    pub const fn column_extend(&self) -> u16 {
        self.column_extend
    }

    /// Returns the zero-based column index. (Java `getColumnIndex()`)
    #[must_use]
    pub const fn column_index(&self) -> u16 {
        self.column_index
    }
}

impl WriteHandler for LoopMergeStrategy {
    fn order(&self) -> i32 {
        // Matches `OrderConstant.FILL_STYLE` — fill-style strategies run last.
        50_000
    }
}

impl AbstractMergeStrategy for LoopMergeStrategy {
    fn merge(
        &mut self,
        _sheet_name: &str,
        _cell: &WriteCellContext,
        _extra: Option<&CellExtra>,
        _relative_row_index: Option<i32>,
    ) {
        // `rust_xlsxwriter` is told to merge the range at write time via
        // `worksheet.merge_range(...)`. The template fill and main writer
        // paths consult this struct to discover the range; the actual
        // mutation happens in those callers.
    }
}