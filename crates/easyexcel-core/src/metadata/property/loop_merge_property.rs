//! Mirrors Java `com.alibaba.excel.metadata.property.LoopMergeProperty`.

/// Mirrors Java `LoopMergeProperty`. (Java `eachRow: int`, `columnExtend: int`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LoopMergeProperty {
    /// Each row. (Java `eachRow`)
    pub each_row: u32,
    /// Extend column. (Java `columnExtend`)
    pub column_extend: u16,
}

impl LoopMergeProperty {
    /// Creates a `LoopMergeProperty`. (Java constructor)
    #[must_use]
    pub const fn new(each_row: u32, column_extend: u16) -> Self {
        Self {
            each_row,
            column_extend,
        }
    }
    /// Returns `eachRow`. (Java `getEachRow()`)
    #[must_use]
    pub const fn each_row(&self) -> u32 {
        self.each_row
    }
    /// Returns `columnExtend`. (Java `getColumnExtend()`)
    #[must_use]
    pub const fn column_extend(&self) -> u16 {
        self.column_extend
    }
}
