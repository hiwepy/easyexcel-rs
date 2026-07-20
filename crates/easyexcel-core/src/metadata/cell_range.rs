//! Mirrors Java `com.alibaba.excel.metadata.CellRange`.

/// Inclusive rectangular cell range.
///
/// Rust port of Java `CellRange`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    /// First row index. (Java `firstRow`)
    pub first_row: i32,
    /// Last row index. (Java `lastRow`)
    pub last_row: i32,
    /// First column index. (Java `firstCol`)
    pub first_col: i32,
    /// Last column index. (Java `lastCol`)
    pub last_col: i32,
}

impl CellRange {
    /// Creates a cell range. (Java constructor)
    #[must_use]
    pub const fn new(first_row: i32, last_row: i32, first_col: i32, last_col: i32) -> Self {
        Self {
            first_row,
            last_row,
            first_col,
            last_col,
        }
    }

    /// Returns the first row index. (Java `getFirstRow()`)
    #[must_use]
    pub const fn first_row(&self) -> i32 {
        self.first_row
    }

    /// Returns the last row index. (Java `getLastRow()`)
    #[must_use]
    pub const fn last_row(&self) -> i32 {
        self.last_row
    }

    /// Returns the first column index. (Java `getFirstCol()`)
    #[must_use]
    pub const fn first_col(&self) -> i32 {
        self.first_col
    }

    /// Returns the last column index. (Java `getLastCol()`)
    #[must_use]
    pub const fn last_col(&self) -> i32 {
        self.last_col
    }
}
