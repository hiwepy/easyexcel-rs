//! Mirrors Java `com.alibaba.excel.read.metadata.ReadSheet`.

/// Mirrors Java `ReadSheet extends ReadBasicParameter`.
///
/// Rust keeps the sheet identity fields used by SAX executors and
/// [`super::AnalysisContextImpl::current_sheet`]. Parameter fields from
/// Java `ReadBasicParameter` remain on [`crate::ReadOptions`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReadSheet {
    /// Zero-based sheet index. (Java `ReadSheet.sheetNo`)
    sheet_no: usize,
    /// Worksheet name. (Java `ReadSheet.sheetName`)
    sheet_name: String,
}

impl ReadSheet {
    /// Mirrors Java `ReadSheet(Integer sheetNo)`.
    #[must_use]
    pub fn new(sheet_no: usize) -> Self {
        Self {
            sheet_no,
            sheet_name: String::new(),
        }
    }

    /// Mirrors Java `ReadSheet(Integer sheetNo, String sheetName)`.
    #[must_use]
    pub fn with_name(sheet_no: usize, sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_no,
            sheet_name: sheet_name.into(),
        }
    }

    /// Returns the zero-based sheet index. (Java `getSheetNo()`)
    #[must_use]
    pub const fn sheet_no(&self) -> usize {
        self.sheet_no
    }

    /// Returns the worksheet name. (Java `getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
}
