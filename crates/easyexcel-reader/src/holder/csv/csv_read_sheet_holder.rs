//! Mirrors Java `com.alibaba.excel.read.metadata.holder.csv.CsvReadSheetHolder`.

use crate::holder::read_sheet_holder::ReadSheetHolder;

/// Mirrors Java `CsvReadSheetHolder extends ReadSheetHolder`.
pub struct CsvReadSheetHolder {
    inner: ReadSheetHolder,
}

impl CsvReadSheetHolder {
    /// Mirrors Java constructor.
    pub fn new(sheet_no: i32, sheet_name: impl Into<String>) -> Self {
        Self { inner: ReadSheetHolder::new(sheet_no, sheet_name) }
    }
    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadSheetHolder { &self.inner }
}
