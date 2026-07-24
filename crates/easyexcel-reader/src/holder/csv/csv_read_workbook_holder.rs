//! Mirrors Java `com.alibaba.excel.read.metadata.holder.csv.CsvReadWorkbookHolder`.

use crate::holder::read_workbook_holder::ReadWorkbookHolder;

/// Mirrors Java `CsvReadWorkbookHolder extends ReadWorkbookHolder`.
#[derive(Debug, Clone)]
pub struct CsvReadWorkbookHolder {
    inner: ReadWorkbookHolder,
}

impl CsvReadWorkbookHolder {
    /// Mirrors Java constructor.
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

impl Default for CsvReadWorkbookHolder {
    fn default() -> Self {
        Self::new()
    }
}
