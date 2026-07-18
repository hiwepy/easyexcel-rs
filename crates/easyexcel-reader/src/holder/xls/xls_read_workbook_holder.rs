//! Mirrors Java `com.alibaba.excel.read.metadata.holder.xls.XlsReadWorkbookHolder`.

use crate::holder::read_workbook_holder::ReadWorkbookHolder;

/// Mirrors Java `XlsReadWorkbookHolder extends ReadWorkbookHolder`.
pub struct XlsReadWorkbookHolder {
    inner: ReadWorkbookHolder,
}

impl XlsReadWorkbookHolder {
    /// Mirrors Java constructor.
    pub fn new() -> Self { Self { inner: ReadWorkbookHolder::default() } }
    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadWorkbookHolder { &self.inner }
}

impl Default for XlsReadWorkbookHolder {
    fn default() -> Self { Self::new() }
}
