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

impl Clone for ExcelError {
    fn clone(&self) -> Self {
        match self {
            Self::Data {
                sheet,
                row,
                column,
                field,
                value,
                message,
            } => Self::Data {
                sheet: sheet.clone(),
                row: *row,
                column: *column,
                field,
                value: value.clone(),
                message: message.clone(),
            },
            Self::SheetNotFound(s) => Self::SheetNotFound(s.clone()),
            Self::Format(s) => Self::Format(s.clone()),
            Self::Unsupported(s) => Self::Unsupported(s.clone()),
            Self::Io(e) => Self::Io(std::io::Error::new(e.kind(), e.to_string())),
        }
    }
}

impl PartialEq for ExcelError {
    fn eq(&self, other: &Self) -> bool {
        match (self, other) {
            (
                Self::Data {
                    sheet: s1,
                    row: r1,
                    column: c1,
                    field: f1,
                    value: v1,
                    message: m1,
                },
                Self::Data {
                    sheet: s2,
                    row: r2,
                    column: c2,
                    field: f2,
                    value: v2,
                    message: m2,
                },
            ) => s1 == s2 && r1 == r2 && c1 == c2 && f1 == f2 && v1 == v2 && m1 == m2,
            (Self::SheetNotFound(a), Self::SheetNotFound(b)) => a == b,
            (Self::Format(a), Self::Format(b)) => a == b,
            (Self::Unsupported(a), Self::Unsupported(b)) => a == b,
            (Self::Io(a), Self::Io(b)) => a.kind() == b.kind() && a.to_string() == b.to_string(),
            _ => false,
        }
    }
}

impl Eq for ExcelError {}
