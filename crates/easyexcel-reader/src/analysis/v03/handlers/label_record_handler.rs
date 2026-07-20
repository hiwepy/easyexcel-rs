//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.LabelRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// Decoded inline-label cell produced by [`LabelRecordHandler`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LabelCell {
    /// Zero-based row. (Java `LabelRecord.getRow`)
    pub row: u32,
    /// Zero-based column. (Java `LabelRecord.getColumn`)
    pub column: usize,
    /// Label text (already trimmed when `auto_trim` was set).
    pub value: String,
}

/// Mirrors Java `LabelRecordHandler`.
#[derive(Debug, Default)]
pub struct LabelRecordHandler;

impl LabelRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Java `LabelRecordHandler.processRecord`.
    #[must_use]
    pub fn process_label(row: u32, column: usize, value: &str, auto_trim: bool) -> LabelCell {
        let value = if auto_trim {
            value.trim().to_owned()
        } else {
            value.to_owned()
        };
        LabelCell {
            row,
            column,
            value,
        }
    }
}

/// BIFF `Label` record sid. (POI `LabelRecord.sid`)
pub const LABEL_SID: u16 = 0x0204;

impl XlsRecordHandler for LabelRecordHandler {
    /// Java `LabelRecordHandler.processRecord` — parses coordinates; string body
    /// decoding is left to a higher-level BIFF reader / [`Self::process_label`].
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != LABEL_SID || data.len() < 8 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        let _ = Self::process_label(row, column, "", false);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn process_label_trims_when_requested() {
        let cell = LabelRecordHandler::process_label(1, 2, " a ", true);
        assert_eq!(cell.value, "a");
    }
}
