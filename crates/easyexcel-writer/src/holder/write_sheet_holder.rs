//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteSheetHolder`.

use crate::MirroredWriteTableHolder as WriteTableHolder;

/// Mirrors Java `WriteSheetHolder extends AbstractWriteHolder`.
///
/// Java's holder stores a POI `Sheet` instance plus the in-flight row
/// cursors. The Rust port reuses [`crate::ExcelWriter`] for the live
/// `rust_xlsxwriter::Worksheet`; this struct is provided for parity so
/// handler context builders can hold an `&WriteSheetHolder`.
pub struct WriteSheetHolder<'a> {
    sheet_name: String,
    sheet_no: i32,
    tables: Vec<WriteTableHolder<'a>>,
    last_row_index: i32,
    has_data: bool,
}

impl<'a> WriteSheetHolder<'a> {
    /// Creates a sheet holder matching the Java `WriteSheetHolder(WriteSheet, WriteWorkbookHolder)` initialiser.
    #[must_use]
    pub fn new(sheet_name: impl Into<String>, sheet_no: i32) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            sheet_no,
            tables: Vec::new(),
            last_row_index: 0,
            has_data: false,
        }
    }

    /// Returns the sheet name. (Java `getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Returns the zero-based sheet index. (Java `getSheetNo()`)
    #[must_use]
    pub const fn sheet_no(&self) -> i32 {
        self.sheet_no
    }

    /// Returns the per-table holders. (Java `getHasBeenInitializedTable()`)
    #[must_use]
    pub fn tables(&self) -> &[WriteTableHolder<'a>] {
        &self.tables
    }

    /// Returns a mutable handle on the per-table holders.
    pub fn tables_mut(&mut self) -> &mut Vec<WriteTableHolder<'a>> {
        &mut self.tables
    }

    /// Returns the last row index. (Java `getLastRowIndex()`)
    #[must_use]
    pub const fn last_row_index(&self) -> i32 {
        self.last_row_index
    }

    /// Returns whether at least one row has been written. (Java `getHasData()`)
    #[must_use]
    pub const fn has_data(&self) -> bool {
        self.has_data
    }

    /// Records the next row index. (Java `getNewRowIndexAndStartDoWrite()` step)
    pub fn advance_row(&mut self) -> i32 {
        self.has_data = true;
        self.last_row_index += 1;
        self.last_row_index
    }
}