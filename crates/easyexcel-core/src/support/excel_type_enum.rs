//! Mirrors Java `com.alibaba.excel.support.ExcelTypeEnum`.
//!
//! Java distinguishes three Excel types by file extension and magic bytes:
//! `XLSX` (`PK\x03\x04`), `XLS` (`D0CF11E0A1B11AE1`), and `CSV` (no magic).
//! Rust mirrors the same three variants.

/// Mirrors Java `ExcelTypeEnum`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelTypeEnum {
    /// CSV format. (Java `CSV`)
    Csv,
    /// Legacy XLS (BIFF) format. (Java `XLS`)
    Xls,
    /// XLSX (OOXML) format. (Java `XLSX`)
    Xlsx,
}

impl ExcelTypeEnum {
    /// Returns the file extension. (Java `getValue()`)
    #[must_use]
    pub const fn value(self) -> &'static str {
        match self {
            Self::Csv => ".csv",
            Self::Xls => ".xls",
            Self::Xlsx => ".xlsx",
        }
    }

    /// Sniffs the type from magic bytes. (Java `recognitionExcelType(InputStream)`)
    #[must_use]
    pub fn from_magic(bytes: &[u8]) -> Self {
        // XLSX magic: 50 4B 03 04
        if bytes.len() >= 4 && bytes[0..4] == [0x50, 0x4B, 0x03, 0x04] {
            return Self::Xlsx;
        }
        // XLS magic: D0 CF 11 E0 A1 B1 1A E1
        if bytes.len() >= 8 && bytes[0..8] == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
            return Self::Xls;
        }
        // CSV has no fixed prefix; default to CSV.
        Self::Csv
    }

    /// Sniffs the type from a file extension.
    #[must_use]
    pub fn from_extension(extension: &str) -> Option<Self> {
        match extension.to_ascii_lowercase().as_str() {
            "csv" => Some(Self::Csv),
            "xls" => Some(Self::Xls),
            "xlsx" => Some(Self::Xlsx),
            _ => None,
        }
    }
}

impl Default for ExcelTypeEnum {
    fn default() -> Self {
        Self::Xlsx
    }
}
