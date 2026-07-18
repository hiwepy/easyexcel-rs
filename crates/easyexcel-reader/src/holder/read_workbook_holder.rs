//! Mirrors Java `com.alibaba.excel.read.metadata.holder.ReadWorkbookHolder`.

/// Mirrors Java `ReadWorkbookHolder extends AbstractReadHolder`.
///
/// Java carries 17 fields. Rust collapses them into the `ReadOptions`
/// struct that already lives in the reader facade. This struct exists
/// for 1:1 API parity.
#[derive(Debug, Clone, Default)]
pub struct ReadWorkbookHolder {
    /// Mirrors `ReadWorkbookHolder.charset`.
    pub charset: easyexcel_core::CsvCharset,
    /// Mirrors `ReadWorkbookHolder.autoCloseStream`.
    pub auto_close_stream: bool,
    /// Mirrors `ReadWorkbookHolder.ignoreEmptyRow`.
    pub ignore_empty_row: bool,
    /// Mirrors `ReadWorkbookHolder.password`.
    pub password: Option<String>,
}
