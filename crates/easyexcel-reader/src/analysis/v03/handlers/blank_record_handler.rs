//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.BlankRecordHandler`.
//!
//! XLS BIFF decoding is owned by `calamine::Xls` today; these helpers encode
//! the Java `processRecord` semantics so a future `XlsSaxAnalyser` can call
//! them without re-deriving the rules.

use super::super::xls_record_handler::XlsRecordHandler;

/// Decoded blank-cell placement produced by [`BlankRecordHandler`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BlankCell {
    /// Zero-based row. (Java `BlankRecord.getRow`)
    pub row: u32,
    /// Zero-based column. (Java `BlankRecord.getColumn`)
    pub column: usize,
}

/// Mirrors Java `BlankRecordHandler`.
#[derive(Debug, Default)]
pub struct BlankRecordHandler {
    /// Most recently decoded blank cell.
    pub last_cell: Option<BlankCell>,
}

impl BlankRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `BlankRecordHandler.processRecord` — emit an empty cell at `(row, column)`.
    #[must_use]
    pub fn process_blank(row: u32, column: usize) -> BlankCell {
        BlankCell { row, column }
    }
}

/// BIFF `Blank` record sid. (POI `BlankRecord.sid`)
pub const BLANK_SID: u16 = 0x0201;

impl XlsRecordHandler for BlankRecordHandler {
    /// Java `BlankRecordHandler.processRecord` — parses `row|col|xf`.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != BLANK_SID || data.len() < 6 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        self.last_cell = Some(Self::process_blank(row, column));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn process_blank_keeps_coordinates() {
        assert_eq!(
            BlankRecordHandler::process_blank(2, 5),
            BlankCell { row: 2, column: 5 }
        );
    }
}
