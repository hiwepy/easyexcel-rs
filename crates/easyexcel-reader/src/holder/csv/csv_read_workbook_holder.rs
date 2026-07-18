//! Mirrors Java `com.alibaba.excel.read.metadata.holder.csv.CsvReadWorkbookHolder`.

use crate::holder::read_workbook_holder::ReadWorkbookHolder;

/// Mirrors Java `CsvReadWorkbookHolder extends ReadWorkbookHolder`.
pub struct CsvReadWorkbookHolder {
    inner: ReadWorkbookHolder,
}

impl CsvReadWorkbookHolder {
    /// Mirrors Java constructor.
    pub fn new() -> Self { Self { inner: ReadWorkbookHolder::default() } }
    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadWorkbookHolder { &self.inner }
}

impl Default for CsvReadWorkbookHolder {
    fn default() -> Self { Self::new() }
}
