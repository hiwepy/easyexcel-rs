//! Mirrors Java `com.alibaba.excel.write.metadata.WriteTable`.

use crate::WriteOptions;
use crate::metadata::WriteBasicParameter;

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
    /// Nullable table-level overrides used for Java parent-holder inheritance.
    pub parameter: WriteBasicParameter,
}

impl WriteTable {
    /// Creates a `WriteTable` with table no 0. (Java `new WriteTable()`)
    #[must_use]
    pub fn new() -> Self {
        Self {
            table_no: 0,
            options: WriteOptions::default(),
            parameter: WriteBasicParameter::default(),
        }
    }

    /// Creates a `WriteTable` with the given table no. (Java `WriteTable.tableNo` setter)
    #[must_use]
    pub fn with_table_no(table_no: i32) -> Self {
        Self {
            table_no,
            options: WriteOptions::default(),
            parameter: WriteBasicParameter::default(),
        }
    }

    /// Returns the zero-based table index. (Java `getTableNo()`)
    #[must_use]
    pub const fn table_no(&self) -> i32 {
        self.table_no
    }

    /// Sets the zero-based table index. (Java `setTableNo(Integer)`)
    pub fn set_table_no(&mut self, table_no: i32) -> &mut Self {
        self.table_no = table_no;
        self
    }

    /// Returns the shared write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }

    /// Returns nullable table-level overrides before parent inheritance.
    #[must_use]
    pub const fn parameter(&self) -> &WriteBasicParameter {
        &self.parameter
    }
}

impl Default for WriteTable {
    fn default() -> Self {
        Self::new()
    }
}
