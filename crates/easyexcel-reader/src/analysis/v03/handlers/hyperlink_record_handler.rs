//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.HyperlinkRecordHandler`.

use easyexcel_core::{CellExtra, CellExtraType};

use super::super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `HyperlinkRecordHandler`.
#[derive(Debug, Default)]
pub struct HyperlinkRecordHandler {
    /// Whether hyperlink extras are enabled. (Java `support`)
    pub enabled: bool,
    /// Last parsed hyperlink extra.
    pub last_extra: Option<CellExtra>,
}

impl HyperlinkRecordHandler {
    /// Creates a handler; `enabled` mirrors Java `support(XlsReadContext)`.
    #[must_use]
    pub fn new(enabled: bool) -> Self {
        Self {
            enabled,
            last_extra: None,
        }
    }

    /// Java `HyperlinkRecordHandler.processRecord`.
    pub fn process_hyperlink(
        &mut self,
        address: Option<String>,
        first_row: u32,
        last_row: u32,
        first_column: usize,
        last_column: usize,
    ) {
        if !self.enabled {
            return;
        }
        self.last_extra = Some(CellExtra::new(
            CellExtraType::Hyperlink,
            address,
            first_row,
            last_row,
            first_column,
            last_column,
        ));
    }
}

impl XlsRecordHandler for HyperlinkRecordHandler {
    fn support(&self) -> bool {
        self.enabled
    }

    /// Java `HyperlinkRecordHandler.processRecord` — sid/range gate; address via helper.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        /// BIFF `Hyperlink` sid (POI `HyperlinkRecord.sid`)
        const HYPERLINK_SID: u16 = 0x01B8;
        if !self.enabled || record_sid != HYPERLINK_SID || data.len() < 8 {
            return;
        }
        let first_row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let last_row = u16::from_le_bytes([data[2], data[3]]) as u32;
        let first_column = u16::from_le_bytes([data[4], data[5]]) as usize;
        let last_column = u16::from_le_bytes([data[6], data[7]]) as usize;
        self.process_hyperlink(None, first_row, last_row, first_column, last_column);
    }
}
