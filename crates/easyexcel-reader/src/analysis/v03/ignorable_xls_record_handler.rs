//! Mirrors Java `com.alibaba.excel.analysis.v03.IgnorableXlsRecordHandler`.

use super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `IgnorableXlsRecordHandler extends XlsRecordHandler`.
///
/// Java marks handlers whose records can be safely ignored.
#[allow(dead_code)]
pub struct IgnorableXlsRecordHandler;

impl XlsRecordHandler for IgnorableXlsRecordHandler {}
