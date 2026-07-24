//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.NumberRecordHandler`.
//!
//! XLS BIFF decoding is owned by `calamine::Xls` today; these helpers encode
//! the Java `processRecord` numeric-cell semantics for a future SAX path.

use super::super::xls_record_handler::XlsRecordHandler;

/// Decoded number cell produced by [`NumberRecordHandler`].
#[derive(Debug, Clone, PartialEq)]
pub struct NumberCell {
    /// Zero-based row. (Java `NumberRecord.getRow`)
    pub row: u32,
    /// Zero-based column. (Java `NumberRecord.getColumn`)
    pub column: usize,
    /// Raw IEEE value. (Java `NumberRecord.getValue`)
    pub value: f64,
    /// Format index from the format-tracking listener (may be 0 when unknown).
    pub format_index: u16,
}

/// Mirrors Java `NumberRecordHandler`.
#[derive(Debug, Default)]
pub struct NumberRecordHandler {
    /// Most recently decoded number cell.
    pub last_cell: Option<NumberCell>,
}

impl NumberRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `NumberRecordHandler.processRecord` (without BuiltinFormats lookup).
    #[must_use]
    pub fn process_number(row: u32, column: usize, value: f64, format_index: u16) -> NumberCell {
        NumberCell {
            row,
            column,
            value,
            format_index,
        }
    }
}

/// BIFF `Number` record sid. (POI `NumberRecord.sid`)
pub const NUMBER_SID: u16 = 0x0203;

impl XlsRecordHandler for NumberRecordHandler {
    /// Java `NumberRecordHandler.processRecord` — parses BIFF Number body
    /// (`row|col|xf|f64`). Formatting lookup stays in [`Self::process_number`].
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != NUMBER_SID || data.len() < 14 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        let mut bits = [0u8; 8];
        bits.copy_from_slice(&data[6..14]);
        let value = f64::from_le_bytes(bits);
        let format_index = u16::from_le_bytes([data[4], data[5]]);
        self.last_cell = Some(Self::process_number(row, column, value, format_index));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn process_number_keeps_value() {
        let cell = NumberRecordHandler::process_number(1, 2, 3.5, 0);
        assert_eq!(cell.row, 1);
        assert_eq!(cell.column, 2);
        assert_eq!(cell.value, 3.5);
    }
}
