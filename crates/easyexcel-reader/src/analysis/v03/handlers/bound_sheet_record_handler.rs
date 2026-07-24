//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.BoundSheetRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// Collected bound-sheet entry (name + BOF position).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BoundSheetEntry {
    /// Sheet display name. (Java `BoundSheetRecord.getSheetname`)
    pub name: String,
    /// Absolute BOF file position used for ordering.
    pub bof_position: u32,
}

/// Mirrors Java `BoundSheetRecordHandler`.
#[derive(Debug, Default)]
pub struct BoundSheetRecordHandler {
    /// Accumulated bound-sheet records. (Java workbook holder list)
    pub sheets: Vec<BoundSheetEntry>,
}

impl BoundSheetRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `BoundSheetRecordHandler.processRecord`.
    pub fn process_bound_sheet(&mut self, name: String, bof_position: u32) {
        self.sheets.push(BoundSheetEntry { name, bof_position });
    }

    /// Java `BoundSheetRecord.orderByBofPosition` — sort by BOF offset ascending.
    pub fn ordered_sheets(&self) -> Vec<BoundSheetEntry> {
        let mut sheets = self.sheets.clone();
        sheets.sort_by_key(|entry| entry.bof_position);
        sheets
    }
}

/// BIFF `BoundSheet` record sid. (POI `BoundSheetRecord.sid`)
pub const BOUND_SHEET_SID: u16 = 0x0085;

impl XlsRecordHandler for BoundSheetRecordHandler {
    /// Java `BoundSheetRecordHandler.processRecord` — reads BOF position and
    /// the BIFF8 short-Unicode sheet name.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != BOUND_SHEET_SID || data.len() < 8 {
            return;
        }
        let bof_position = u32::from_le_bytes([data[0], data[1], data[2], data[3]]);
        let character_count = data[6] as usize;
        let is_utf16 = data[7] & 0x01 != 0;
        let raw = &data[8..];
        let name = if is_utf16 {
            let byte_count = character_count.saturating_mul(2);
            if raw.len() < byte_count {
                return;
            }
            let units = raw[..byte_count]
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect::<Vec<_>>();
            String::from_utf16_lossy(&units)
        } else {
            if raw.len() < character_count {
                return;
            }
            raw[..character_count]
                .iter()
                .map(|byte| char::from(*byte))
                .collect()
        };
        self.process_bound_sheet(name, bof_position);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn orders_by_bof_position() {
        let mut handler = BoundSheetRecordHandler::new();
        handler.process_bound_sheet("B".into(), 200);
        handler.process_bound_sheet("A".into(), 100);
        let ordered = handler.ordered_sheets();
        assert_eq!(ordered[0].name, "A");
        assert_eq!(ordered[1].name, "B");
    }

    #[test]
    fn decodes_compressed_biff8_sheet_name() {
        let mut handler = BoundSheetRecordHandler::new();
        let mut payload = vec![0x20, 0, 0, 0, 0, 0, 4, 0];
        payload.extend_from_slice(b"Data");
        handler.process_record(BOUND_SHEET_SID, &payload);
        assert_eq!(handler.sheets[0].name, "Data");
        assert_eq!(handler.sheets[0].bof_position, 32);
    }
}
