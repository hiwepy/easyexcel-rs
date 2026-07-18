//! Mirrors Java `com.alibaba.excel.metadata.property.OnceAbsoluteMergeProperty`.

/// Mirrors Java `OnceAbsoluteMergeProperty`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OnceAbsoluteMergeProperty {
    /// First row index. (Java `firstRowIndex`)
    pub first_row_index: i32,
    /// Last row index. (Java `lastRowIndex`)
    pub last_row_index: i32,
    /// First column index. (Java `firstColumnIndex`)
    pub first_column_index: i32,
    /// Last column index. (Java `lastColumnIndex`)
    pub last_column_index: i32,
}

impl OnceAbsoluteMergeProperty {
    /// Creates a `OnceAbsoluteMergeProperty`. (Java constructor)
    #[allow(clippy::too_many_arguments)]
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
}
