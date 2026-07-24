//! Mirrors Java `com.alibaba.excel.write.merge.LoopMergeStrategy`.

use easyexcel_core::{
    CellExtra, ExcelError, LoopMergeProperty, Result, WriteCellContext, WriteHandler,
};

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
    ///
    /// # Errors
    ///
    /// Returns an error when `each_rows < 1`, `column_extend < 1`, or when
    /// Java's combined constraint `eachRow < 2 && columnExtend < 2` holds.
    pub fn new(each_rows: u32, column_extend: u16, column_index: u16) -> Result<Self> {
        // Java: eachRow < 1 → IllegalArgumentException("EachRows must be greater than 1")
        if each_rows < 1 {
            return Err(ExcelError::Format(
                "EachRows must be greater than 1".to_owned(),
            ));
        }
        // Java: columnExtend < 1 → IllegalArgumentException("ColumnExtend must be greater than 1")
        if column_extend < 1 {
            return Err(ExcelError::Format(
                "ColumnExtend must be greater than 1".to_owned(),
            ));
        }
        // Java: eachRow < 2 && columnExtend < 2 → IllegalArgumentException(
        //   "EachRows or ColumnExtend cannot be less than 2, otherwise they will not be merged")
        if each_rows < 2 && column_extend < 2 {
            return Err(ExcelError::Format(
                "EachRows or ColumnExtend cannot be less than 2, otherwise they will not be merged"
                    .to_owned(),
            ));
        }
        Ok(Self {
            each_rows,
            column_extend,
            column_index,
        })
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
        // Java `LoopMergeStrategy` does not override `order()`.
        easyexcel_core::constant::order_constant::DEFAULT_ORDER
    }

    fn style_loop_merge(&self) -> Option<(LoopMergeProperty, usize)> {
        Some((
            LoopMergeProperty::new(self.each_rows, self.column_extend),
            usize::from(self.column_index),
        ))
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
