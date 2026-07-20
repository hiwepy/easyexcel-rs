//! Mirrors Java `com.alibaba.excel.read.builder.ExcelReaderSheetBuilder`.
//!
//! Java signature (10 methods + AbstractExcelReaderParameterBuilder
//! inherited methods):
//! ```java
//! public class ExcelReaderSheetBuilder
//!     extends AbstractExcelReaderParameterBuilder<ExcelReaderSheetBuilder, ReadSheet> {
//!     public ExcelReaderSheetBuilder();
//!     public ExcelReaderSheetBuilder(ExcelReader excelReader);
//!     public ExcelReaderSheetBuilder sheetNo(Integer sheetNo);
//!     public ExcelReaderSheetBuilder sheetName(String sheetName);
//!     public ReadSheet build();
//!     public void doRead();
//!     public <T> List<T> doReadSync();
//!     protected ReadSheet parameter();
//!     // inherited from AbstractExcelReaderParameterBuilder:
//!     public T headRowNumber(Integer headRowNumber);
//!     public T useScientificFormat(Boolean useScientificFormat);
//!     public T registerReadListener(ReadListener<?> readListener);
//! }
//! ```

use easyexcel_core::ReadListener;

use crate::context::read_sheet::ReadSheet;
use crate::excel_reader::ExcelReader;

/// Mirrors Java `ExcelReaderSheetBuilder extends AbstractExcelReaderParameterBuilder`.
///
/// Java has a back-reference `private ExcelReader excelReader` that
/// `doRead()` and `doReadSync()` use to invoke the read. Rust mirrors
/// this through `with_excel_reader`; `do_read` and `do_read_sync`
/// remain on the reader-side facade (`EasyExcel::read`) so this
/// builder only carries the static metadata.
#[derive(Debug, Clone, Default)]
pub struct ExcelReaderSheetBuilder {
    /// Mirrors `ExcelReaderSheetBuilder.sheetNo`.
    pub sheet_no: Option<i32>,
    /// Mirrors `ExcelReaderSheetBuilder.sheetName`.
    pub sheet_name: Option<String>,
    /// Mirrors `AbstractExcelReaderParameterBuilder.headRowNumber`.
    pub head_row_number: Option<i32>,
    /// Mirrors `AbstractExcelReaderParameterBuilder.useScientificFormat`.
    pub use_scientific_format: Option<bool>,
}

impl ExcelReaderSheetBuilder {
    /// Creates an empty sheet builder. (Java `ExcelReaderSheetBuilder()`)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates a sheet builder bound to an [`ExcelReader`].
    /// (Java `ExcelReaderSheetBuilder(ExcelReader excelReader)`)
    ///
    /// The back-reference is currently stored in the typed
    /// `EasyExcel::read` / `ExcelReader::sheet` facades, not on
    /// this builder, so the constructor only forwards the call.
    #[must_use]
    pub fn with_excel_reader<T, L>(_excel_reader: &ExcelReader<T, L>) -> Self
    where
        T: easyexcel_core::ExcelRow,
        L: ReadListener<T>,
    {
        Self::default()
    }

    /// Sets the zero-based sheet index. (Java `sheetNo(Integer)`)
    #[must_use]
    pub fn sheet_no(mut self, sheet_no: i32) -> Self {
        self.sheet_no = Some(sheet_no);
        self
    }

    /// Sets the sheet name. (Java `sheetName(String)`)
    #[must_use]
    pub fn sheet_name(mut self, sheet_name: impl Into<String>) -> Self {
        self.sheet_name = Some(sheet_name.into());
        self
    }

    /// Returns the typed `ReadSheet` view used by the reader.
    /// (Java `protected ReadSheet parameter()`)
    #[must_use]
    pub fn parameter(&self) -> ReadSheet {
        self.build()
    }

    /// Returns the finalised [`ReadSheet`] for use with the
    /// [`ExcelReader`]. (Java `public ReadSheet build()`)
    ///
    /// Sheet_no defaults to 0 (matching Java's `new ReadSheet()`).
    #[must_use]
    pub fn build(&self) -> ReadSheet {
        let no = self.sheet_no.unwrap_or(0).max(0) as usize;
        if let Some(name) = &self.sheet_name {
            ReadSheet::with_name(no, name.clone())
        } else {
            ReadSheet::new(no)
        }
    }

    /// Sets the head row number (zero-based). (Java
    /// `AbstractExcelReaderParameterBuilder.headRowNumber(Integer)`)
    #[must_use]
    pub fn head_row_number(mut self, head_row_number: i32) -> Self {
        self.head_row_number = Some(head_row_number);
        self
    }

    /// Toggles scientific-format coercion for numeric cells.
    /// (Java `AbstractExcelReaderParameterBuilder.useScientificFormat(Boolean)`)
    #[must_use]
    pub fn use_scientific_format(mut self, enabled: bool) -> Self {
        self.use_scientific_format = Some(enabled);
        self
    }
}