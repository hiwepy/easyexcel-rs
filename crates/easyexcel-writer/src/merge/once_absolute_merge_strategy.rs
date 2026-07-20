//! Mirrors Java `com.alibaba.excel.write.merge.OnceAbsoluteMergeStrategy`.

use easyexcel_core::{
    CellExtra, OnceAbsoluteMergeProperty, WriteCellContext, WriteHandler,
};

use crate::merge::abstract_merge_strategy::AbstractMergeStrategy;

/// Mirrors Java `OnceAbsoluteMergeStrategy implements SheetWriteHandler`.
///
/// Registered instances are consumed by the XLSX write path via
/// [`WriteHandler::style_once_absolute_merge`] (in addition to type-level
/// `@OnceAbsoluteMerge` metadata).
pub struct OnceAbsoluteMergeStrategy {
    first_row_index: i32,
    last_row_index: i32,
    first_column_index: i32,
    last_column_index: i32,
}

impl OnceAbsoluteMergeStrategy {
    /// Creates the strategy. (Java
    /// `OnceAbsoluteMergeStrategy(int, int, int, int)`)
    ///
    /// Java throws when any index is negative; Rust stores the values and the
    /// write path skips invalid regions (same as annotation apply).
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

    /// Creates from annotation/runtime property.
    /// (Java `OnceAbsoluteMergeStrategy(OnceAbsoluteMergeProperty)`)
    #[must_use]
    pub const fn from_property(property: OnceAbsoluteMergeProperty) -> Self {
        Self::new(
            property.first_row_index,
            property.last_row_index,
            property.first_column_index,
            property.last_column_index,
        )
    }

    /// Returns the merge region as a property. (Java getters)
    #[must_use]
    pub const fn to_property(&self) -> OnceAbsoluteMergeProperty {
        OnceAbsoluteMergeProperty::new(
            self.first_row_index,
            self.last_row_index,
            self.first_column_index,
            self.last_column_index,
        )
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

    fn style_once_absolute_merge(&self) -> Option<OnceAbsoluteMergeProperty> {
        // Java `afterSheetCreate` → `addMergedRegionUnsafe`
        Some(self.to_property())
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
        // Absolute merges run once at sheet create via
        // `WriteHandler::style_once_absolute_merge`, not per cell.
    }
}
