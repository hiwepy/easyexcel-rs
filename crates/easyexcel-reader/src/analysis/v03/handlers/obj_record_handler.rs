//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.ObjRecordHandler`.
//!
//! Tracks the current drawing/object id used later by note/text handlers.

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `ObjRecordHandler`.
#[derive(Debug, Default)]
pub struct ObjRecordHandler {
    /// Last seen object / shape id. (Java sheet holder)
    pub temp_object_index: Option<u32>,
}

impl ObjRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `ObjRecordHandler.processRecord`.
    pub fn process_obj(&mut self, object_id: u32) {
        self.temp_object_index = Some(object_id);
    }
}

/// BIFF `Obj` record sid. (POI `ObjRecord.sid`)
pub const OBJ_SID: u16 = 0x005D;

impl XlsRecordHandler for ObjRecordHandler {
    /// Java `ObjRecordHandler.processRecord` — sid gate; object id via [`Self::process_obj`].
    fn process_record(&mut self, record_sid: u16, _data: &[u8]) {
        if record_sid == OBJ_SID {
            // Full CommonObjectDataSubRecord parse is deferred; keep last index.
            let _ = self.temp_object_index;
        }
    }
}
