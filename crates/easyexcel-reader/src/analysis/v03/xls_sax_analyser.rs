//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsSaxAnalyser`.

/// Mirrors Java `XlsSaxAnalyser implements HSSFListener, ExcelReadExecutor`.
///
/// The Java side registers 19 record handlers in a lookup map keyed
/// by `Record.sid`. Rust delegates XLS parsing to `calamine::Xls`.
#[allow(dead_code)]
pub struct XlsSaxAnalyser;
