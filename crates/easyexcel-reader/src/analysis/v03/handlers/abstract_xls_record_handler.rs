//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.AbstractXlsRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `AbstractXlsRecordHandler implements XlsRecordHandler`.
///
/// Java base class provides default `support() == true` and leaves
/// `processRecord` abstract; concrete handlers override it.
#[derive(Debug, Default)]
pub struct AbstractXlsRecordHandler;

impl AbstractXlsRecordHandler {
    /// Creates the abstract base (rarely constructed on its own).
    #[must_use]
    pub fn new() -> Self {
        Self
    }
}

impl XlsRecordHandler for AbstractXlsRecordHandler {
    fn support(&self) -> bool {
        true
    }

    fn process_record(&mut self, _record_sid: u16, _data: &[u8]) {}
}
