//! Mirrors Java `com.alibaba.excel.read.metadata.holder.xlsx.XlsxReadWorkbookHolder`.

use crate::holder::read_workbook_holder::ReadWorkbookHolder;

/// Mirrors Java `XlsxReadWorkbookHolder extends ReadWorkbookHolder`.
pub struct XlsxReadWorkbookHolder {
    inner: ReadWorkbookHolder,
}

impl XlsxReadWorkbookHolder {
    /// Mirrors Java `XlsxReadWorkbookHolder(ReadWorkbook)`.
    pub fn new() -> Self {
        Self { inner: ReadWorkbookHolder::default() }
    }
    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadWorkbookHolder { &self.inner }
}

impl Default for XlsxReadWorkbookHolder {
    fn default() -> Self { Self::new() }
}
