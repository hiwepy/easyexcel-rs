//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.TextObjectRecordHandler`.
//!
//! Stores comment text under the current object id for later `NoteRecord` use.

use std::collections::HashMap;

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `TextObjectRecordHandler`.
#[derive(Debug, Default)]
pub struct TextObjectRecordHandler {
    /// shapeId → comment text. (Java `objectCacheMap`)
    pub object_cache: HashMap<u32, String>,
}

impl TextObjectRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `TextObjectRecordHandler.processRecord`.
    pub fn process_text(&mut self, object_id: u32, text: String) {
        self.object_cache.insert(object_id, text);
    }

    /// Lookup used by [`super::note_record_handler::NoteRecordHandler`].
    #[must_use]
    pub fn get(&self, object_id: u32) -> Option<&str> {
        self.object_cache.get(&object_id).map(String::as_str)
    }
}

/// BIFF `TextObject` record sid. (POI `TextObjectRecord.sid`)
pub const TEXT_OBJECT_SID: u16 = 0x01B6;

impl XlsRecordHandler for TextObjectRecordHandler {
    /// Java `TextObjectRecordHandler.processRecord` — parses TxO (0x01B6)
    /// to extract shapeId + linked Continue record text.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != TEXT_OBJECT_SID && record_sid != CONTINUE_SID {
            return;
        }
        if record_sid == TEXT_OBJECT_SID && data.len() >= 4 {
            let object_id = u16::from_le_bytes([data[2], data[3]]) as u32;
            // Extract text from TxO record data starting at byte 10
            // (after grbit[2] + shapeId[2] + reserved[8] = 12 bytes min)
            if data.len() > 12 {
                let text_bytes = &data[12..];
                // Simple byte to string (ISO-8859-1 / ASCII for BIFF8 TxO)
                let text: String = text_bytes
                    .iter()
                    .take_while(|&&b| b != 0)
                    .map(|&b| b as char)
                    .collect();
                if !text.is_empty() {
                    self.object_cache.insert(object_id, text);
                    return;
                }
            }
            self.object_cache
                .entry(object_id)
                .or_insert_with(|| format!("TxO_{object_id}"));
        }
        // Handle CONTINUE record (0x003C) — text continuation
        if record_sid == CONTINUE_SID && data.len() >= 2 {
            let text: String = data
                .iter()
                .take_while(|&&b| b != 0)
                .map(|&b| b as char)
                .collect();
            if !text.is_empty() && !self.object_cache.is_empty() {
                // Attach to the most recent TxO entry
                if let Some((_, val)) = self.object_cache.iter_mut().last() {
                    val.push_str(&text);
                }
            }
        }
    }
}

/// BIFF `Continue` record sid.
const CONTINUE_SID: u16 = 0x003C;
