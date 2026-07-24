//! Mirrors Java `com.alibaba.excel.read.builder.ExcelReaderTableBuilder`.
//!
//! Java signature (5 members):
//! ```java
//! public class ExcelReaderTableBuilder
//!     extends AbstractExcelReaderParameterBuilder<ExcelReaderTableBuilder, ReadTable> {
//!     private ReadTable readTable;
//!     public ExcelReaderTableBuilder();
//!     public ExcelReaderTableBuilder(ExcelReader excelReader);
//!     public ExcelReaderTableBuilder tableNo(Integer tableNo);
//!     public ReadTable build();
//!     protected ReadTable parameter();
//! }
//! ```

use easyexcel_core::ReadListener;

use crate::excel_reader::ExcelReader;
use crate::metadata::read_table::ReadTable;

/// Mirrors Java `ExcelReaderTableBuilder extends AbstractExcelReaderParameterBuilder`.
///
/// Rust: table-level configuration is sparse in this port because
/// `ReadTable` is an in-memory struct (the Java type itself is
/// minimal). The builder here mostly carries `head_row_number` and
/// `use_scientific_format` for parity with the sheet builder.
#[derive(Debug, Clone, Default)]
pub struct ExcelReaderTableBuilder {
    /// Mirrors `ExcelReaderTableBuilder.tableNo`.
    pub table_no: Option<i32>,
    /// Mirrors `AbstractExcelReaderParameterBuilder.headRowNumber`.
    pub head_row_number: Option<i32>,
    /// Mirrors `AbstractExcelReaderParameterBuilder.useScientificFormat`.
    pub use_scientific_format: Option<bool>,
}

impl ExcelReaderTableBuilder {
    /// Creates an empty table builder. (Java `ExcelReaderTableBuilder()`)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates a table builder bound to an [`ExcelReader`].
    /// (Java `ExcelReaderTableBuilder(ExcelReader)`)
    #[must_use]
    pub fn with_excel_reader<T, L>(_excel_reader: &ExcelReader<T, L>) -> Self
    where
        T: easyexcel_core::ExcelRow,
        L: ReadListener<T>,
    {
        Self::default()
    }

    /// Sets the zero-based table index. (Java `tableNo(Integer)`)
    #[must_use]
    pub fn table_no(mut self, table_no: i32) -> Self {
        self.table_no = Some(table_no);
        self
    }

    /// Returns the typed `ReadTable` view used by the reader.
    /// (Java `protected ReadTable parameter()`)
    #[must_use]
    pub fn parameter(&self) -> ReadTable {
        self.build()
    }

    /// Builds the underlying table configuration. (Java `ReadTable build()`)
    ///
    /// Rust port: returns a `ReadTable` carrying the configured
    /// `table_no`. Callers compose this with `ExcelReader::table(...)`.
    #[must_use]
    pub fn build(&self) -> ReadTable {
        ReadTable::with_table_no(self.table_no.unwrap_or(0))
    }

    /// Sets the head row number. (Java
    /// `AbstractExcelReaderParameterBuilder.headRowNumber(Integer)`)
    #[must_use]
    pub fn head_row_number(mut self, head_row_number: i32) -> Self {
        self.head_row_number = Some(head_row_number);
        self
    }

    /// Toggles scientific-format coercion. (Java
    /// `AbstractExcelReaderParameterBuilder.useScientificFormat(Boolean)`)
    #[must_use]
    pub fn use_scientific_format(mut self, enabled: bool) -> Self {
        self.use_scientific_format = Some(enabled);
        self
    }
}
