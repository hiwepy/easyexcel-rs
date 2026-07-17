//! Mirrors the union of Java `com.alibaba.excel.exception.*` classes.

use thiserror::Error;

/// All public easyexcel errors with row and column diagnostics where applicable.
///
/// Java uses seven `RuntimeException` subclasses (`ExcelCommonException`,
/// `ExcelAnalysisException`, `ExcelAnalysisStopException`, etc.). Rust
/// collapses them into a single `Error` enum with `thiserror` for
/// ergonomic `Display` / `From<io::Error>` integration.
#[derive(Debug, Error)]
pub enum ExcelError {
    /// A cell-to-field conversion error. (Java `ExcelDataConvertException`)
    #[error(
        "sheet={sheet}, row={row}, column={column:?}, field={field}, value={value:?}: {message}"
    )]
    Data {
        /// Sheet name.
        sheet: String,
        /// Zero-based row index.
        row: u32,
        /// Zero-based column index.
        column: Option<usize>,
        /// Rust field name.
        field: &'static str,
        /// Original cell text.
        value: String,
        /// Human-readable failure reason.
        message: String,
    },
    /// A requested worksheet does not exist. (Java `SheetNotFoundException`)
    #[error("worksheet not found: {0}")]
    SheetNotFound(String),
    /// The workbook or OOXML package is invalid. (Java `ExcelAnalysisException`)
    #[error("excel format error: {0}")]
    Format(String),
    /// The requested operation is not supported by the selected engine. (Java `ExcelCommonException`)
    #[error("unsupported operation: {0}")]
    Unsupported(String),
    /// An I/O operation failed. (Java `ExcelCommonException` wrapping `IOException`)
    #[error(transparent)]
    Io(#[from] std::io::Error),
}
