//! Mirrors Java `com.alibaba.excel.read.builder.ExcelReaderSheetBuilder`.

/// Mirrors Java `ExcelReaderSheetBuilder extends AbstractExcelReaderParameterBuilder`.
#[derive(Debug, Clone, Default)]
pub struct ExcelReaderSheetBuilder {
    /// Mirrors `ExcelReaderSheetBuilder.sheetNo`.
    pub sheet_no: Option<i32>,
    /// Mirrors `ExcelReaderSheetBuilder.sheetName`.
    pub sheet_name: Option<String>,
}

impl ExcelReaderSheetBuilder {
    /// Creates a builder.
    #[must_use]
    pub fn new() -> Self { Self::default() }

    /// Sets the zero-based sheet index. (Java `sheetNo(Integer)`)
    pub fn sheet_no(mut self, sheet_no: i32) -> Self {
        self.sheet_no = Some(sheet_no);
        self
    }

    /// Sets the sheet name. (Java `sheetName(String)`)
    pub fn sheet_name(mut self, sheet_name: impl Into<String>) -> Self {
        self.sheet_name = Some(sheet_name.into());
        self
    }
}
