//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.StringRecordHandler`.
//!
//! Completes a pending string-formula cell created by [`super::formula_record_handler`].

use super::super::xls_record_handler::XlsRecordHandler;
use super::formula_record_handler::{FormulaCell, FormulaRecordHandler};
use crate::analysis::v03::biff_string::decode_unicode_string_segments;

/// Mirrors Java `StringRecordHandler`.
#[derive(Debug, Default)]
pub struct StringRecordHandler {
    /// Most recently decoded formula string result.
    pub last_value: Option<String>,
}

impl StringRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
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

    /// Decodes one logical BIFF8 String record and its CONTINUE bodies.
    pub fn decode_segments(segments: &[Vec<u8>]) -> easyexcel_core::Result<String> {
        decode_unicode_string_segments(segments)
    }

    /// Stores an already decoded String record value.
    pub fn process_decoded(&mut self, value: String) {
        self.last_value = Some(value);
    }
}

/// BIFF `String` record sid (formula result). (POI `StringRecord.sid`)
pub const STRING_SID: u16 = 0x0207;

impl XlsRecordHandler for StringRecordHandler {
    /// Java `StringRecordHandler.processRecord` — sid gate; pair with
    /// [`Self::process_string`] and a live [`FormulaRecordHandler`].
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != STRING_SID {
            return;
        }
        if let Ok(value) = Self::decode_segments(&[data.to_vec()]) {
            self.process_decoded(value);
        }
    }
}
