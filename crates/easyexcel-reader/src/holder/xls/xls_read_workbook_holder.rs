//! Mirrors Java `com.alibaba.excel.read.metadata.holder.xls.XlsReadWorkbookHolder`.

use crate::holder::read_workbook_holder::ReadWorkbookHolder;

/// Mirrors Java `XlsReadWorkbookHolder extends ReadWorkbookHolder`.
#[derive(Debug, Clone)]
pub struct XlsReadWorkbookHolder {
    inner: ReadWorkbookHolder,
    need_read_sheet: bool,
}

impl XlsReadWorkbookHolder {
    /// Mirrors Java constructor.
    pub fn new() -> Self {
        Self {
            inner: ReadWorkbookHolder::default(),
            need_read_sheet: true,
        }
    }

    /// Creates the format-specific holder from resolved workbook options.
    #[must_use]
    pub fn from_options(options: &crate::ReadOptions) -> Self {
        Self {
            inner: ReadWorkbookHolder::from_options(options),
            need_read_sheet: true,
        }
    }

    /// Returns the inner holder.
    pub const fn inner(&self) -> &ReadWorkbookHolder {
        &self.inner
    }

    /// Returns mutable common workbook state.
    pub const fn inner_mut(&mut self) -> &mut ReadWorkbookHolder {
        &mut self.inner
    }

    /// Returns whether the main record pass should process worksheet data.
    #[must_use]
    pub const fn need_read_sheet(&self) -> bool {
        self.need_read_sheet
    }

    /// Controls worksheet-data processing.
    ///
    /// Java `XlsListSheetListener` disables it during its metadata-only pass.
    pub const fn set_need_read_sheet(&mut self, need_read_sheet: bool) {
        self.need_read_sheet = need_read_sheet;
    }
}

impl Default for XlsReadWorkbookHolder {
    fn default() -> Self {
        Self::new()
    }
}
