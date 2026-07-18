//! Mirrors Java `com.alibaba.excel.write.metadata.WriteTable`.

use crate::WriteOptions;

/// Mirrors Java `WriteTable extends WriteBasicParameter`.
///
/// Java carries a `tableNo` field. Rust reuses [`WriteOptions`] for the
/// common base and adds the table index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteTable {
    /// Mirrors `WriteTable.tableNo`.
    pub table_no: i32,
    /// Mirrors the remaining `WriteBasicParameter` fields.
    pub options: WriteOptions,
}

impl WriteTable {
    /// Creates a `WriteTable` with table no 0. (Java `new WriteTable()`)
    #[must_use]
    pub fn new() -> Self {
        Self {
            table_no: 0,
            options: WriteOptions::default(),
        }
    }

    /// Creates a `WriteTable` with the given table no. (Java `WriteTable.tableNo` setter)
    #[must_use]
    pub fn with_table_no(table_no: i32) -> Self {
        Self {
            table_no,
            options: WriteOptions::default(),
        }
    }

    /// Returns the zero-based table index. (Java `getTableNo()`)
    #[must_use]
    pub const fn table_no(&self) -> i32 {
        self.table_no
    }

    /// Returns the shared write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }
}

impl Default for WriteTable {
    fn default() -> Self {
        Self::new()
    }
}