//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.NoteRecordHandler`.

use easyexcel_core::{CellExtra, CellExtraType};

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `NoteRecordHandler` (comment / note).
#[derive(Debug, Default)]
pub struct NoteRecordHandler {
    /// Whether comment extras are enabled. (Java `support`)
    pub enabled: bool,
    /// Last parsed comment extra.
    pub last_extra: Option<CellExtra>,
}

impl NoteRecordHandler {
    /// Creates a handler; `enabled` mirrors Java `support(XlsReadContext)`.
    #[must_use]
    pub fn new(enabled: bool) -> Self {
        Self {
            enabled,
            last_extra: None,
        }
    }

    /// Java `NoteRecordHandler.processRecord`.
    ///
    /// `text` comes from `objectCacheMap.get(shapeId)` in Java.
    pub fn process_note(&mut self, text: Option<String>, row: u32, column: usize) {
        if !self.enabled {
            return;
        }
        self.last_extra = Some(CellExtra::new(
            CellExtraType::Comment,
            text,
            row,
            row,
            column,
            column,
        ));
    }
}

impl XlsRecordHandler for NoteRecordHandler {
    fn support(&self) -> bool {
        self.enabled
    }

    /// Java `NoteRecordHandler.processRecord` — parses row/col; text via cache.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        /// BIFF `Note` sid (POI `NoteRecord.sid`)
        const NOTE_SID: u16 = 0x001C;
        if !self.enabled || record_sid != NOTE_SID || data.len() < 6 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        self.process_note(None, row, column);
    }
}
