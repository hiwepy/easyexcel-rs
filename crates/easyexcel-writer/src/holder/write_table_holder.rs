//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteTableHolder`.

/// Mirrors Java `WriteTableHolder extends AbstractWriteHolder`.
///
/// Java's holder carries a POI `Sheet` plus a `tableNo` field. The Rust port
/// mirrors the type so the [`crate::ExcelWriterTableBuilder`] can return a
/// `WriteTableHolder` for parity.
pub struct WriteTableHolder<'a> {
    table_no: i32,
    parent_sheet: Option<&'a str>,
    last_row_index: i32,
}

impl<'a> WriteTableHolder<'a> {
    /// Creates a table holder matching the Java `WriteTableHolder(WriteTable, WriteSheetHolder)` initialiser.
    #[must_use]
    pub fn new(table_no: i32) -> Self {
        Self {
            table_no,
            parent_sheet: None,
            last_row_index: 0,
        }
    }

    /// Returns the parent sheet name, if any. (Java `getParentWriteSheetHolder().getSheetName()`)
    #[must_use]
    pub fn parent_sheet(&self) -> Option<&str> {
        self.parent_sheet
    }

    /// Sets the parent sheet name.
    pub fn set_parent_sheet(&mut self, parent: &'a str) {
        self.parent_sheet = Some(parent);
    }

    /// Returns the zero-based table index. (Java `getTableNo()`)
    #[must_use]
    pub const fn table_no(&self) -> i32 {
        self.table_no
    }

    /// Returns the last row index. (Java `getLastRowIndex()`)
    #[must_use]
    pub const fn last_row_index(&self) -> i32 {
        self.last_row_index
    }
}