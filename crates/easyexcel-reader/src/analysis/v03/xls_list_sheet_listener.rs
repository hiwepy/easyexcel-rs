//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsListSheetListener`.

/// Mirrors Java `XlsListSheetListener implements HSSFListener`.
///
/// Java's listener pre-scans BIFF records to enumerate sheet names
/// before the main read. Rust uses `calamine::Xls::worksheet_names()`
/// directly.
#[allow(dead_code)]
pub struct XlsListSheetListener;
