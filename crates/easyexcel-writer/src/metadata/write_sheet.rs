//! Mirrors Java `com.alibaba.excel.write.metadata.WriteSheet`.

use crate::WriteOptions;
use crate::metadata::WriteBasicParameter;

/// Mirrors Java `WriteSheet extends WriteBasicParameter`.
///
/// Java stores `sheetNo` and `sheetName`. Rust reuses [`WriteOptions`] and
/// extends the type with the two fields so 1:1 naming is preserved.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheet {
    /// Mirrors `WriteSheet.sheetNo`.
    pub sheet_no: i32,
    /// Mirrors `WriteSheet.sheetName`.
    pub sheet_name: String,
    /// Mirrors the remaining `WriteBasicParameter` fields.
    pub options: WriteOptions,
    /// Nullable sheet-level overrides before workbook inheritance.
    pub parameter: WriteBasicParameter,
}

impl WriteSheet {
    /// Creates a `WriteSheet` matching Java `new WriteSheet()`.
    #[must_use]
    pub fn new() -> Self {
        Self {
            sheet_no: 0,
            sheet_name: String::new(),
            options: WriteOptions::default(),
            parameter: WriteBasicParameter::default(),
        }
    }

    /// Creates a `WriteSheet` with the given sheet no. (Java `WriteSheet(sheetNo)`)
    #[must_use]
    pub fn with_sheet_no(sheet_no: i32) -> Self {
        Self {
            sheet_no,
            sheet_name: String::new(),
            options: WriteOptions::default(),
            parameter: WriteBasicParameter::default(),
        }
    }

    /// Creates a `WriteSheet` with the given sheet no and name. (Java `WriteSheet(sheetNo, sheetName)`)
    #[must_use]
    pub fn with_sheet(sheet_no: i32, sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_no,
            sheet_name: sheet_name.into(),
            options: WriteOptions::default(),
            parameter: WriteBasicParameter::default(),
        }
    }

    /// Returns the zero-based sheet index. (Java `getSheetNo()`)
    #[must_use]
    pub const fn sheet_no(&self) -> i32 {
        self.sheet_no
    }

    /// Sets the zero-based sheet index. (Java `setSheetNo(Integer)`)
    pub fn set_sheet_no(&mut self, sheet_no: i32) -> &mut Self {
        self.sheet_no = sheet_no;
        self
    }

    /// Returns the sheet name. (Java `getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Sets the sheet name. (Java `setSheetName(String)`)
    pub fn set_sheet_name(&mut self, sheet_name: impl Into<String>) -> &mut Self {
        self.sheet_name = sheet_name.into();
        self
    }

    /// Returns the shared write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }

    /// Returns nullable sheet-level overrides before workbook inheritance.
    #[must_use]
    pub const fn parameter(&self) -> &WriteBasicParameter {
        &self.parameter
    }
}

impl Default for WriteSheet {
    fn default() -> Self {
        Self::new()
    }
}
