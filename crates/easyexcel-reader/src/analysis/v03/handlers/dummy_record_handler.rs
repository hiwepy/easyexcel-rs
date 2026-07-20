//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.DummyRecordHandler`.
//!
//! Handles POI "dummy" records that mark end-of-row and missing cells.

use std::collections::HashMap;

use super::super::xls_record_handler::XlsRecordHandler;
use super::blank_record_handler::BlankCell;

/// Events synthesised by [`DummyRecordHandler`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DummyRecordEvent {
    /// Java `LastCellOfRowDummyRecord` — flush the current row.
    EndRow {
        /// Zero-based row index to emit.
        row: u32,
    },
    /// Java `MissingCellDummyRecord` — insert empty if absent.
    MissingCell(BlankCell),
}

/// Mirrors Java `DummyRecordHandler`.
#[derive(Debug, Default)]
pub struct DummyRecordHandler;

impl DummyRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Java `LastCellOfRowDummyRecord` branch.
    #[must_use]
    pub fn process_last_cell_of_row(row: u32) -> DummyRecordEvent {
        DummyRecordEvent::EndRow { row }
    }

    /// Java `MissingCellDummyRecord` branch — `putIfAbsent` semantics.
    ///
    /// Returns `Some(MissingCell)` only when the column is not already present
    /// (see EasyExcel issue #2236).
    pub fn process_missing_cell(
        row: u32,
        column: usize,
        existing: &HashMap<usize, ()>,
    ) -> Option<DummyRecordEvent> {
        if existing.contains_key(&column) {
            return None;
        }
        Some(DummyRecordEvent::MissingCell(BlankCell { row, column }))
    }
}

impl XlsRecordHandler for DummyRecordHandler {
    /// Java `DummyRecordHandler.processRecord` — POI synthesised dummy records
    /// are not true BIFF sids; use [`Self::process_last_cell_of_row`] /
    /// [`Self::process_missing_cell`].
    fn process_record(&mut self, _record_sid: u16, _data: &[u8]) {
        // No-op by design (matches Java's instanceof branches on DummyRecord).
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn missing_cell_skips_existing_columns() {
        let mut map = HashMap::new();
        map.insert(1usize, ());
        assert!(DummyRecordHandler::process_missing_cell(0, 1, &map).is_none());
        assert!(DummyRecordHandler::process_missing_cell(0, 2, &map).is_some());
    }
}
