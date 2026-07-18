//! Mirrors Java `com.alibaba.excel.write.merge.OnceAbsoluteMergeStrategy`.

use easyexcel_core::{CellExtra, WriteCellContext, WriteHandler};

use crate::merge::abstract_merge_strategy::AbstractMergeStrategy;

/// Mirrors Java `OnceAbsoluteMergeStrategy implements SheetWriteHandler`.
pub struct OnceAbsoluteMergeStrategy {
    first_row_index: i32,
    last_row_index: i32,
    first_column_index: i32,
    last_column_index: i32,
}

impl OnceAbsoluteMergeStrategy {
    /// Creates the strategy. (Java
    /// `OnceAbsoluteMergeStrategy(int, int, int, int)`)
    #[must_use]
    pub const fn new(
        first_row_index: i32,
        last_row_index: i32,
        first_column_index: i32,
        last_column_index: i32,
    ) -> Self {
        Self {
            first_row_index,
            last_row_index,
            first_column_index,
            last_column_index,
        }
    }

    /// Returns the first row index. (Java `getFirstRowIndex()`)
    #[must_use]
    pub const fn first_row_index(&self) -> i32 {
        self.first_row_index
    }

    /// Returns the last row index. (Java `getLastRowIndex()`)
    #[must_use]
    pub const fn last_row_index(&self) -> i32 {
        self.last_row_index
    }

    /// Returns the first column index. (Java `getFirstColumnIndex()`)
    #[must_use]
    pub const fn first_column_index(&self) -> i32 {
        self.first_column_index
    }

    /// Returns the last column index. (Java `getLastColumnIndex()`)
    #[must_use]
    pub const fn last_column_index(&self) -> i32 {
        self.last_column_index
    }
}

impl WriteHandler for OnceAbsoluteMergeStrategy {
    fn order(&self) -> i32 {
        -60_000
    }
}

impl AbstractMergeStrategy for OnceAbsoluteMergeStrategy {
    fn merge(
        &mut self,
        _sheet_name: &str,
        _cell: &WriteCellContext,
        _extra: Option<&CellExtra>,
        _relative_row_index: Option<i32>,
    ) {
        // The actual `worksheet.merge_range` call is performed by the
        // `ExcelTemplateWriter` / `ExcelWriter` after collecting all
        // strategy instances.
    }
}