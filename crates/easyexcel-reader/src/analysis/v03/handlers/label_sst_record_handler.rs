//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.LabelSstRecordHandler`.
//!
//! Resolves an SST index through a caller-supplied cache lookup, matching
//! Java's `ReadCache.get(sstIndex)` path.

use super::super::xls_record_handler::XlsRecordHandler;

/// Outcome of [`LabelSstRecordHandler::process_label_sst`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LabelSstCell {
    /// Empty cell when the cache is missing or the index is absent.
    Empty {
        /// Zero-based row.
        row: u32,
        /// Zero-based column.
        column: usize,
    },
    /// Resolved shared-string cell.
    String {
        /// Zero-based row.
        row: u32,
        /// Zero-based column.
        column: usize,
        /// Resolved text (already trimmed when `auto_trim` was set).
        value: String,
    },
}

/// Mirrors Java `LabelSstRecordHandler`.
#[derive(Debug, Default)]
pub struct LabelSstRecordHandler;

impl LabelSstRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Java `LabelSstRecordHandler.processRecord`.
    ///
    /// `resolve` maps SST index → string (`ReadCache.get`); `None` yields empty.
    pub fn process_label_sst(
        row: u32,
        column: usize,
        sst_index: usize,
        auto_trim: bool,
        resolve: &dyn Fn(usize) -> Option<String>,
    ) -> LabelSstCell {
        match resolve(sst_index) {
            None => LabelSstCell::Empty { row, column },
            Some(mut data) => {
                if auto_trim {
                    data = data.trim().to_owned();
                }
                LabelSstCell::String {
                    row,
                    column,
                    value: data,
                }
            }
        }
    }
}

/// BIFF `LabelSST` record sid. (POI `LabelSSTRecord.sid`)
pub const LABEL_SST_SID: u16 = 0x00FD;

impl XlsRecordHandler for LabelSstRecordHandler {
    /// Java `LabelSstRecordHandler.processRecord` — accepts LabelSST sid and
    /// validates the 10-byte BIFF body (`row|col|xf|sstIndex`). Full cache
    /// resolution uses [`LabelSstRecordHandler::process_label_sst`].
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != LABEL_SST_SID || data.len() < 10 {
            return;
        }
        let _row = u16::from_le_bytes([data[0], data[1]]);
        let _col = u16::from_le_bytes([data[2], data[3]]);
        let _sst = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_sst_yields_empty() {
        let cell = LabelSstRecordHandler::process_label_sst(0, 1, 9, false, &|_| None);
        assert_eq!(cell, LabelSstCell::Empty { row: 0, column: 1 });
    }

    #[test]
    fn auto_trim_strips_whitespace() {
        let cell =
            LabelSstRecordHandler::process_label_sst(0, 0, 0, true, &|_| Some("  hi  ".into()));
        assert_eq!(
            cell,
            LabelSstCell::String {
                row: 0,
                column: 0,
                value: "hi".into()
            }
        );
    }
}
