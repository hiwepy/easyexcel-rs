//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.MergeCellsRecordHandler`.

use easyexcel_core::{CellExtra, CellExtraType};

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `MergeCellsRecordHandler`.
#[derive(Debug, Default)]
pub struct MergeCellsRecordHandler {
    /// Whether merge extras are enabled. (Java `support`)
    pub enabled: bool,
    /// Last emitted merge extras from one record (may contain multiple areas).
    pub last_extras: Vec<CellExtra>,
}

impl MergeCellsRecordHandler {
    /// Creates a handler; `enabled` mirrors Java `support(XlsReadContext)`.
    #[must_use]
    pub fn new(enabled: bool) -> Self {
        Self {
            enabled,
            last_extras: Vec::new(),
        }
    }

    /// Java `MergeCellsRecordHandler.support`.
    #[must_use]
    pub fn support(&self) -> bool {
        self.enabled
    }

    /// Java `MergeCellsRecordHandler.processRecord` for one merged area.
    pub fn process_area(
        &mut self,
        first_row: u32,
        last_row: u32,
        first_column: usize,
        last_column: usize,
    ) {
        if !self.enabled {
            return;
        }
        self.last_extras.push(CellExtra::new(
            CellExtraType::Merge,
            None,
            first_row,
            last_row,
            first_column,
            last_column,
        ));
    }

    /// Drains extras accumulated for the current record.
    pub fn take_extras(&mut self) -> Vec<CellExtra> {
        std::mem::take(&mut self.last_extras)
    }
}

impl XlsRecordHandler for MergeCellsRecordHandler {
    fn support(&self) -> bool {
        self.enabled
    }

    /// Java `MergeCellsRecordHandler.processRecord` — parses area count + ranges.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        /// BIFF `MergeCells` sid (POI `MergeCellsRecord.sid`)
        const MERGE_CELLS_SID: u16 = 0x00E5;
        if !self.enabled || record_sid != MERGE_CELLS_SID || data.len() < 2 {
            return;
        }
        self.last_extras.clear();
        let count = u16::from_le_bytes([data[0], data[1]]) as usize;
        let mut offset = 2;
        for _ in 0..count {
            if offset + 8 > data.len() {
                break;
            }
            let first_row = u16::from_le_bytes([data[offset], data[offset + 1]]) as u32;
            let last_row = u16::from_le_bytes([data[offset + 2], data[offset + 3]]) as u32;
            let first_column = u16::from_le_bytes([data[offset + 4], data[offset + 5]]) as usize;
            let last_column = u16::from_le_bytes([data[offset + 6], data[offset + 7]]) as usize;
            self.process_area(first_row, last_row, first_column, last_column);
            offset += 8;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn disabled_handler_ignores_areas() {
        let mut handler = MergeCellsRecordHandler::new(false);
        handler.process_area(0, 1, 0, 1);
        assert!(handler.take_extras().is_empty());
    }

    #[test]
    fn enabled_handler_collects_areas() {
        let mut handler = MergeCellsRecordHandler::new(true);
        handler.process_area(0, 1, 0, 2);
        let extras = handler.take_extras();
        assert_eq!(extras.len(), 1);
        assert_eq!(extras[0].extra_type(), CellExtraType::Merge);
    }
}
