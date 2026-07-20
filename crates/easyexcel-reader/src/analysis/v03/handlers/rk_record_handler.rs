//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.RkRecordHandler`.
//!
//! Note: Java oddly materialises an *empty* cell for RK records (historical
//! EasyExcel behaviour). We mirror that exactly.

use super::super::xls_record_handler::XlsRecordHandler;
use super::blank_record_handler::BlankCell;

/// Mirrors Java `RkRecordHandler`.
#[derive(Debug, Default)]
pub struct RkRecordHandler;

impl RkRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Java `RkRecordHandler.processRecord` — always yields an empty cell.
    #[must_use]
    pub fn process_rk(row: u32, column: usize) -> BlankCell {
        BlankCell { row, column }
    }
}

/// BIFF `RK` record sid. (POI `RKRecord.sid`)
pub const RK_SID: u16 = 0x027E;

impl XlsRecordHandler for RkRecordHandler {
    /// Java `RkRecordHandler.processRecord` — yields empty cell (EasyExcel quirk).
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != RK_SID || data.len() < 4 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        let _ = Self::process_rk(row, column);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn process_rk_is_empty_cell() {
        assert_eq!(
            RkRecordHandler::process_rk(3, 4),
            BlankCell { row: 3, column: 4 }
        );
    }
}
