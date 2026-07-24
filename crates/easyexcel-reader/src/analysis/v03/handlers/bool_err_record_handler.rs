//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.BoolErrRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// Decoded boolean cell produced by [`BoolErrRecordHandler`].
///
/// Java's handler only materialises the boolean branch via
/// `BoolErrRecord.getBooleanValue()` (error branch is not exposed separately).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BoolCell {
    /// Zero-based row.
    pub row: u32,
    /// Zero-based column.
    pub column: usize,
    /// Boolean value. (Java `getBooleanValue`)
    pub value: bool,
}

/// Mirrors Java `BoolErrRecordHandler`.
#[derive(Debug, Default)]
pub struct BoolErrRecordHandler {
    /// Most recently decoded boolean cell.
    pub last_cell: Option<BoolCell>,
}

impl BoolErrRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `BoolErrRecordHandler.processRecord`.
    #[must_use]
    pub fn process_bool(row: u32, column: usize, value: bool) -> BoolCell {
        BoolCell { row, column, value }
    }
}

/// BIFF `BoolErr` record sid. (POI `BoolErrRecord.sid`)
pub const BOOL_ERR_SID: u16 = 0x0205;

impl XlsRecordHandler for BoolErrRecordHandler {
    /// Java `BoolErrRecordHandler.processRecord` — boolean branch only.
    /// Layout: `row|col|xf|value:u8|isError:u8`.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != BOOL_ERR_SID || data.len() < 8 {
            return;
        }
        let is_error = data[7] != 0;
        if is_error {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        let value = data[6] != 0;
        self.last_cell = Some(Self::process_bool(row, column, value));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn process_bool_keeps_flag() {
        assert!(BoolErrRecordHandler::process_bool(0, 0, true).value);
    }
}
