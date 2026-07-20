//! Mirrors Java `com.alibaba.excel.metadata.Cell`.

/// Cell coordinate contract.
///
/// Java `Cell` exposes row and column indices. Rust uses snake_case getters
/// to match project conventions while preserving Java semantics.
///
/// Rust port of Java `Cell`.
pub trait Cell {
    /// Returns the zero-based row index. (Java `getRowIndex()`)
    fn row_index(&self) -> Option<i32>;

    /// Returns the zero-based column index. (Java `getColumnIndex()`)
    fn column_index(&self) -> Option<i32>;
}
