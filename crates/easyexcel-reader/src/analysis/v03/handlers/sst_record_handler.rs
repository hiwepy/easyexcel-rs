//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.SstRecordHandler`.
//!
//! The dispatcher assembles physical CONTINUE records and supplies the decoded
//! strings, matching Java's `XlsCache(SSTRecord)` responsibility.

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `SstRecordHandler`.
#[derive(Debug, Default)]
pub struct SstRecordHandler {
    /// Number of unique strings announced by the SST. (Java `getNumUniqueStrings`)
    pub unique_string_count: Option<u32>,
    /// Decoded shared strings in SST index order. (Java `XlsCache`)
    pub strings: Vec<String>,
}

impl SstRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `SstRecordHandler.processRecord` — bookkeeping only (cache filled elsewhere).
    pub fn process_sst(&mut self, unique_string_count: u32) {
        self.unique_string_count = Some(unique_string_count);
    }

    /// Installs a fully decoded SST after CONTINUE records have been assembled.
    pub fn process_decoded_sst(&mut self, unique_string_count: u32, strings: Vec<String>) {
        self.unique_string_count = Some(unique_string_count);
        self.strings = strings;
    }

    /// Resolves one SST index.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(String::as_str)
    }
}

/// BIFF `SST` record sid. (POI `SSTRecord.sid`)
pub const SST_SID: u16 = 0x00FC;

impl XlsRecordHandler for SstRecordHandler {
    /// Java `SstRecordHandler.processRecord` — reads `cstTotal`/`cstUnique` header.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != SST_SID || data.len() < 8 {
            return;
        }
        let unique = u32::from_le_bytes([data[4], data[5], data[6], data[7]]);
        self.process_sst(unique);
    }
}
