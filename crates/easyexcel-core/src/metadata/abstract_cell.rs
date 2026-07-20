//! Mirrors Java `com.alibaba.excel.metadata.AbstractCell`.

use super::cell::Cell;

/// Base cell coordinate holder.
///
/// Rust port of Java `AbstractCell implements Cell`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct AbstractCell {
    /// Row index. (Java `rowIndex`)
    pub row_index: Option<i32>,
    /// Column index. (Java `columnIndex`)
    pub column_index: Option<i32>,
}

impl AbstractCell {
    /// Creates an empty cell coordinate. (Java default constructor)
    #[must_use]
    pub const fn new() -> Self {
        Self {
            row_index: None,
            column_index: None,
        }
    }

    /// Creates a cell coordinate with explicit indices. (Java setter chain)
    #[must_use]
    pub const fn with_indices(row_index: i32, column_index: i32) -> Self {
        Self {
            row_index: Some(row_index),
            column_index: Some(column_index),
        }
    }
}

impl Cell for AbstractCell {
    fn row_index(&self) -> Option<i32> {
        self.row_index
    }

    fn column_index(&self) -> Option<i32> {
        self.column_index
    }
}
