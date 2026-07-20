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
    /// Java `TextObjectRecordHandler.processRecord` — sid gate; use [`Self::process_text`].
    fn process_record(&mut self, record_sid: u16, _data: &[u8]) {
        let _ = record_sid == TEXT_OBJECT_SID;
    }
}
