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
    /// Whether `sheet_no` was explicitly configured.
    ///
    /// Java stores a nullable `Integer`; Rust keeps the existing numeric getter
    /// while retaining the null-vs-zero distinction needed by `SheetUtils.match`.
    sheet_no_explicit: bool,
    /// Worksheet name. (Java `ReadSheet.sheetName`)
    sheet_name: String,
    /// Sheet-scoped header row count inherited from `ReadBasicParameter`.
    head_row_number: Option<u32>,
    /// Sheet-scoped scientific-format override inherited from `ReadBasicParameter`.
    use_scientific_format: Option<bool>,
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
            sheet_no_explicit: true,
            sheet_name: String::new(),
            head_row_number: None,
            use_scientific_format: None,
        }
    }

    /// Mirrors Java `ReadSheet(Integer sheetNo, String sheetName)`.
    #[must_use]
    pub fn with_name(sheet_no: usize, sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_no,
            sheet_no_explicit: true,
            sheet_name: sheet_name.into(),
            head_row_number: None,
            use_scientific_format: None,
        }
    }

    /// Creates a name-only sheet selector.
    ///
    /// Mirrors `new ReadSheet()` followed by `setSheetName(...)`, leaving the
    /// Java `sheetNo` value null.
    #[must_use]
    pub fn named(sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            ..Self::default()
        }
    }

    /// Returns the zero-based sheet index. (Java `getSheetNo()`)
    #[must_use]
    pub const fn sheet_no(&self) -> usize {
        self.sheet_no
    }

    /// Returns whether a sheet number was explicitly configured.
    #[must_use]
    pub const fn has_sheet_no(&self) -> bool {
        self.sheet_no_explicit
    }

    /// Returns the zero-based sheet index as `i32` (Java boxing). (Java `getSheetNo()`)
    #[must_use]
    pub const fn sheet_no_i32(&self) -> i32 {
        self.sheet_no as i32
    }

    /// Sets the zero-based sheet index. (Java `setSheetNo(Integer)`)
    pub fn set_sheet_no(&mut self, sheet_no: usize) -> &mut Self {
        self.sheet_no = sheet_no;
        self.sheet_no_explicit = true;
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

    /// Returns the sheet-scoped header row count.
    ///
    /// Mirrors `ReadBasicParameter.getHeadRowNumber()`.
    #[must_use]
    pub const fn head_row_number(&self) -> Option<u32> {
        self.head_row_number
    }

    /// Overrides the workbook header row count for this sheet.
    ///
    /// Mirrors `ReadBasicParameter.setHeadRowNumber(Integer)`.
    pub fn set_head_row_number(&mut self, head_row_number: u32) -> &mut Self {
        self.head_row_number = Some(head_row_number);
        self
    }

    /// Returns the sheet-scoped scientific-format override.
    ///
    /// Mirrors `ReadBasicParameter.getUseScientificFormat()`.
    #[must_use]
    pub const fn use_scientific_format(&self) -> Option<bool> {
        self.use_scientific_format
    }

    /// Overrides scientific formatting for this sheet.
    ///
    /// Mirrors `ReadBasicParameter.setUseScientificFormat(Boolean)`.
    pub fn set_use_scientific_format(&mut self, enabled: bool) -> &mut Self {
        self.use_scientific_format = Some(enabled);
        self
    }

    /// Copies common basic-parameter fields from another ReadSheet.
    /// (Java `copyBasicParameter(ReadSheet other)`)
    ///
    pub fn copy_basic_parameter(&mut self, other: &ReadSheet) -> &mut Self {
        self.head_row_number = other.head_row_number;
        self.use_scientific_format = other.use_scientific_format;
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
