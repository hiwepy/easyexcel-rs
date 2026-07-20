//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.IndexRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `IndexRecordHandler`.
#[derive(Debug, Default)]
pub struct IndexRecordHandler {
    /// Approximate total rows from `IndexRecord.getLastRowAdd1`.
    pub approximate_total_row_number: Option<u32>,
}

impl IndexRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `IndexRecordHandler.processRecord`.
    pub fn process_index(&mut self, last_row_add_1: u32) {
        self.approximate_total_row_number = Some(last_row_add_1);
    }
}

/// BIFF `Index` record sid. (POI `IndexRecord.sid`)
pub const INDEX_SID: u16 = 0x020B;

impl XlsRecordHandler for IndexRecordHandler {
    /// Java `IndexRecordHandler.processRecord` — reads `lastRowAdd1` when present.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != INDEX_SID || data.len() < 16 {
            return;
        }
        // IndexRecord: reserved(4) + firstRow(4) + lastRowAdd1(4) + ...
        let last_row_add_1 = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        self.process_index(last_row_add_1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn process_index_stores_total() {
        let mut handler = IndexRecordHandler::new();
        handler.process_index(42);
        assert_eq!(handler.approximate_total_row_number, Some(42));
    }
}
