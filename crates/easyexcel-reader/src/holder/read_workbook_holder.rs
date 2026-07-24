//! Mirrors Java `com.alibaba.excel.read.metadata.holder.ReadWorkbookHolder`.

use crate::context::read_sheet::ReadSheet;

/// Mirrors Java `ReadWorkbookHolder extends AbstractReadHolder`.
///
/// Java carries 17 fields. Rust collapses them into the `ReadOptions`
/// struct that already lives in the reader facade. This struct exists
/// for 1:1 API parity.
#[derive(Debug, Clone, Default)]
pub struct ReadWorkbookHolder {
    /// Mirrors `ReadWorkbookHolder.charset`.
    pub charset: easyexcel_core::CsvCharset,
    /// Mirrors `ReadWorkbookHolder.autoCloseStream`.
    pub auto_close_stream: bool,
    /// Mirrors `ReadWorkbookHolder.ignoreEmptyRow`.
    pub ignore_empty_row: bool,
    /// Mirrors `ReadWorkbookHolder.password`.
    pub password: Option<String>,
    /// Workbooks sheets discovered by the format executor.
    ///
    /// Mirrors `ReadWorkbookHolder.actualSheetDataList`.
    pub actual_sheet_data_list: Option<Vec<ReadSheet>>,
}

impl ReadWorkbookHolder {
    /// Resolves workbook-level holder state from the public read options.
    ///
    /// Mirrors Java `ReadWorkbookHolder(ReadWorkbook, ...)` propagation before
    /// a format-specific context is constructed.
    #[must_use]
    pub fn from_options(options: &crate::ReadOptions) -> Self {
        Self {
            charset: options.charset.clone(),
            auto_close_stream: true,
            ignore_empty_row: options.ignore_empty_row,
            password: options.password.clone(),
            actual_sheet_data_list: None,
        }
    }

    /// Returns format-discovered sheets in workbook order.
    #[must_use]
    pub fn actual_sheet_data_list(&self) -> Option<&[ReadSheet]> {
        self.actual_sheet_data_list.as_deref()
    }

    /// Stores format-discovered sheets.
    pub fn set_actual_sheet_data_list(&mut self, sheets: Vec<ReadSheet>) {
        self.actual_sheet_data_list = Some(sheets);
    }
}
