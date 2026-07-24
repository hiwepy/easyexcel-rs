//! Mirrors Java `com.alibaba.excel.read.metadata.holder.xlsx.XlsxReadWorkbookHolder`.

use crate::holder::read_workbook_holder::ReadWorkbookHolder;

/// Mirrors Java `XlsxReadWorkbookHolder extends ReadWorkbookHolder`.
#[derive(Debug, Clone)]
pub struct XlsxReadWorkbookHolder {
    inner: ReadWorkbookHolder,
}

impl XlsxReadWorkbookHolder {
    /// Mirrors Java `XlsxReadWorkbookHolder(ReadWorkbook)`.
    pub fn new() -> Self {
        Self {
            inner: ReadWorkbookHolder::default(),
        }
    }

    /// Creates the format-specific holder from resolved workbook options.
    #[must_use]
    pub fn from_options(options: &crate::ReadOptions) -> Self {
        Self {
            inner: ReadWorkbookHolder::from_options(options),
        }
    }

    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadWorkbookHolder {
        &self.inner
    }
}

impl Default for XlsxReadWorkbookHolder {
    fn default() -> Self {
        Self::new()
    }
}
