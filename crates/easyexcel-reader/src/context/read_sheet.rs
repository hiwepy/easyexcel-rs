//! Mirrors Java `com.alibaba.excel.read.metadata.ReadSheet`.

use std::fmt;

/// Mirrors Java `ReadSheet extends ReadBasicParameter`.
///
/// Rust keeps the sheet identity fields used by SAX executors and
/// [`super::AnalysisContextImpl::current_sheet`]. Parameter fields from
/// Java `ReadBasicParameter` remain on [`crate::ReadOptions`].
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ReadSheet {
    /// Zero-based sheet index. (Java `ReadSheet.sheetNo`)
    sheet_no: usize,
    /// Worksheet name. (Java `ReadSheet.sheetName`)
    sheet_name: String,
}

impl ReadSheet {
    /// Mirrors Java `ReadSheet()`.
    #[must_use]
    pub fn default_construction() -> Self {
        Self::default()
    }

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

    /// Returns the zero-based sheet index as `i32` (Java boxing). (Java `getSheetNo()`)
    #[must_use]
    pub const fn sheet_no_i32(&self) -> i32 {
        self.sheet_no as i32
    }

    /// Sets the zero-based sheet index. (Java `setSheetNo(Integer)`)
    pub fn set_sheet_no(&mut self, sheet_no: usize) -> &mut Self {
        self.sheet_no = sheet_no;
        self
    }

    /// Returns the worksheet name. (Java `getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Sets the worksheet name. (Java `setSheetName(String)`)
    pub fn set_sheet_name(&mut self, sheet_name: impl Into<String>) -> &mut Self {
        self.sheet_name = sheet_name.into();
        self
    }

    /// Copies common basic-parameter fields from another ReadSheet.
    /// (Java `copyBasicParameter(ReadSheet other)`)
    ///
    /// Rust port: only the identity fields are relevant. The Java
    /// fields `headRowNumber` / `customObject` etc. live on
    /// [`crate::ReadOptions`], not on the sheet metadata.
    pub fn copy_basic_parameter(&mut self, other: &ReadSheet) -> &mut Self {
        self.sheet_no = other.sheet_no;
        self.sheet_name.clone_from(&other.sheet_name);
        self
    }
}

impl fmt::Display for ReadSheet {
    /// Mirrors Java `ReadSheet.toString()`.
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "ReadSheet{{sheetNo={}, sheetName='{}'}}",
            self.sheet_no, self.sheet_name
        )
    }
}