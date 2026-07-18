//! Mirrors Java `com.alibaba.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder`.

use crate::holder::read_sheet_holder::ReadSheetHolder;

/// Mirrors Java `XlsxReadSheetHolder extends ReadSheetHolder`.
pub struct XlsxReadSheetHolder {
    inner: ReadSheetHolder,
}

impl XlsxReadSheetHolder {
    /// Mirrors Java constructor.
    pub fn new(sheet_no: i32, sheet_name: impl Into<String>) -> Self {
        Self { inner: ReadSheetHolder::new(sheet_no, sheet_name) }
    }
    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadSheetHolder { &self.inner }
}
