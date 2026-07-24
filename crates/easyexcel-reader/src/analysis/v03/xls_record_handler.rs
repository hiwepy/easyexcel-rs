//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsRecordHandler`.

/// Handler contract for one or more BIFF record SIDs.
pub trait XlsRecordHandler {
    /// Whether this handler is enabled for the current read configuration.
    fn support(&self) -> bool {
        true
    }

    /// Processes a physical BIFF record body.
    ///
    /// This method is deliberately required: concrete compatibility handlers
    /// cannot silently inherit an empty implementation.
    fn process_record(&mut self, record_sid: u16, data: &[u8]);
}
