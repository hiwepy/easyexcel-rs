//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsListSheetListener`.

use std::path::{Path, PathBuf};

use easyexcel_core::Result;

use crate::context::{DefaultXlsReadContext, ReadSheet};
use crate::{ReadOptions, list_xls_sheets};

/// Mirrors Java `XlsListSheetListener implements HSSFListener`.
///
/// Java's listener pre-scans BIFF records to enumerate sheet names
/// before the main read. Rust performs the same metadata-only pass through
/// calamine and stores the result in `actual_sheet_data_list`.
pub struct XlsListSheetListener<'a> {
    xls_read_context: &'a mut DefaultXlsReadContext,
    path: PathBuf,
    options: ReadOptions,
    sheet_list: Vec<ReadSheet>,
}

impl<'a> XlsListSheetListener<'a> {
    /// Creates the metadata-only listener.
    ///
    /// Mirrors Java `XlsListSheetListener(XlsReadContext)`, including
    /// `needReadSheet = false`.
    pub fn new(
        xls_read_context: &'a mut DefaultXlsReadContext,
        path: impl Into<PathBuf>,
        options: ReadOptions,
    ) -> Self {
        xls_read_context
            .xls_read_workbook_holder_mut()
            .set_need_read_sheet(false);
        Self {
            xls_read_context,
            path: path.into(),
            options,
            sheet_list: Vec::new(),
        }
    }

    /// Executes the real XLS metadata scan and stores discovered sheets.
    ///
    /// # Errors
    ///
    /// Propagates XLS open or metadata parsing errors.
    pub fn execute(&mut self) -> Result<&[ReadSheet]> {
        self.sheet_list = list_xls_sheets(&self.path, &self.options)?
            .into_iter()
            .map(|(sheet_no, sheet_name)| ReadSheet::with_name(sheet_no, sheet_name))
            .collect();
        self.xls_read_context
            .xls_read_workbook_holder_mut()
            .inner_mut()
            .set_actual_sheet_data_list(self.sheet_list.clone());
        Ok(&self.sheet_list)
    }

    /// Returns the bound XLS path.
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Returns the last successfully discovered sheet list.
    #[must_use]
    pub fn sheet_list(&self) -> &[ReadSheet] {
        &self.sheet_list
    }
}
