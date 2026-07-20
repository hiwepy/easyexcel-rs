//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.StringRecordHandler`.
//!
//! Completes a pending string-formula cell created by [`super::formula_record_handler`].

use super::formula_record_handler::{FormulaCell, FormulaRecordHandler};
use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `StringRecordHandler`.
#[derive(Debug, Default)]
pub struct StringRecordHandler;

impl StringRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Java `StringRecordHandler.processRecord` — applies string onto pending formula.
    pub fn process_string(
        formula_handler: &mut FormulaRecordHandler,
        value: String,
        auto_trim: bool,
    ) -> Option<(FormulaCell, String)> {
        let text = if auto_trim {
            value.trim().to_owned()
        } else {
            value
        };
        formula_handler
            .complete_pending_string(text.clone())
            .map(|cell| (cell, text))
    }
}

/// BIFF `String` record sid (formula result). (POI `StringRecord.sid`)
pub const STRING_SID: u16 = 0x0207;

impl XlsRecordHandler for StringRecordHandler {
    /// Java `StringRecordHandler.processRecord` — sid gate; pair with
    /// [`Self::process_string`] and a live [`FormulaRecordHandler`].
    fn process_record(&mut self, record_sid: u16, _data: &[u8]) {
        let _ = record_sid == STRING_SID;
    }
}
