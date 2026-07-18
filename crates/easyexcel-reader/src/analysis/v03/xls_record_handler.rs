//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsRecordHandler`.

/// Mirrors Java `XlsRecordHandler`.
pub trait XlsRecordHandler {
    /// Whether this handler supports the current context.
    fn support(&self) -> bool { true }
    /// Called for each record. (Java `processRecord(XlsReadContext, Record)`)
    fn process_record(&mut self, _record_sid: u16, _data: &[u8]) {}
}
