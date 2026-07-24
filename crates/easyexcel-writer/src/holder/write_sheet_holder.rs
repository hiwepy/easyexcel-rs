//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteSheetHolder`.

use std::ops::{Deref, DerefMut};

use crate::MirroredWriteTableHolder as WriteTableHolder;
use crate::holder::abstract_write_holder::AbstractWriteHolder;
use crate::metadata::WriteBasicParameter;

/// Mirrors Java `WriteSheetHolder extends AbstractWriteHolder`.
///
/// Java's holder stores a POI `Sheet` instance plus the in-flight row
/// cursors. The Rust port reuses [`crate::ExcelWriter`] for the live
/// `rust_xlsxwriter::Worksheet`; this owned builder-side mirror remains for
/// Java package/API parity. Runtime callbacks use
/// [`easyexcel_core::WriteSheetHolderView`] instead of a fake POI sheet.
pub struct WriteSheetHolder<'a> {
    abstract_holder: AbstractWriteHolder,
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
            abstract_holder: AbstractWriteHolder::default(),
            sheet_name: sheet_name.into(),
            sheet_no,
            tables: Vec::new(),
            last_row_index: 0,
            has_data: false,
        }
    }

    /// Creates a sheet holder and resolves nullable values against its parent.
    #[must_use]
    pub fn from_parameter(
        sheet_name: impl Into<String>,
        sheet_no: i32,
        parameter: &WriteBasicParameter,
        parent: &AbstractWriteHolder,
    ) -> Self {
        let mut holder = Self::new(sheet_name, sheet_no);
        holder.abstract_holder = AbstractWriteHolder::from_parameter(parameter, Some(parent));
        holder
    }

    /// Returns the inherited write-holder state.
    #[must_use]
    pub const fn abstract_holder(&self) -> &AbstractWriteHolder {
        &self.abstract_holder
    }

    /// Returns mutable inherited write-holder state.
    pub const fn abstract_holder_mut(&mut self) -> &mut AbstractWriteHolder {
        &mut self.abstract_holder
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

impl Deref for WriteSheetHolder<'_> {
    type Target = AbstractWriteHolder;

    fn deref(&self) -> &Self::Target {
        &self.abstract_holder
    }
}

impl DerefMut for WriteSheetHolder<'_> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.abstract_holder
    }
}
