//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvRow`.
//!
//! Java's CSV metadata classes simulate Excel cell/row/sheet/workbook
//! types for the CSV engine. Rust uses `CellValue` and `csv::Writer`
//! directly, so these are 1:1 type aliases for API parity.

/// Type alias mirroring Java `CsvRow`.
#[allow(dead_code)]
pub type CsvRow = ();
