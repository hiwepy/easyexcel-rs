//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.BofRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// POI `BOFRecord` type codes used by EasyExcel.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BofType {
    /// Workbook-level BOF.
    Workbook,
    /// Worksheet-level BOF.
    Worksheet,
    /// Other (chart, macro, …) — ignored by Java.
    Other,
}

/// Side-effects requested by [`BofRecordHandler`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BofAction {
    /// Reset workbook sheet cursor. (TYPE_WORKBOOK)
    ResetWorkbook,
    /// Ignore non-worksheet BOF.
    Ignore,
    /// Begin / skip a worksheet sheet.
    BeginWorksheet {
        /// Whether the matched sheet should be read (`ignoreRecord = false`).
        read_sheet: bool,
        /// Next `readSheetIndex` after this BOF.
        next_read_sheet_index: usize,
    },
}

/// Mirrors Java `BofRecordHandler`.
#[derive(Debug, Default)]
pub struct BofRecordHandler;

impl BofRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self
    }

    /// Java `BofRecordHandler.processRecord` decision (sheet list already built).
    #[must_use]
    pub fn decide(
        bof_type: BofType,
        read_sheet_index: Option<usize>,
        sheet_matched: bool,
    ) -> BofAction {
        match bof_type {
            BofType::Workbook => BofAction::ResetWorkbook,
            BofType::Other => BofAction::Ignore,
            BofType::Worksheet => {
                let index = read_sheet_index.unwrap_or(0);
                BofAction::BeginWorksheet {
                    read_sheet: sheet_matched,
                    next_read_sheet_index: index.saturating_add(1),
                }
            }
        }
    }
}

/// BIFF `BOF` record sid. (POI `BOFRecord.sid`)
pub const BOF_SID: u16 = 0x0809;

impl XlsRecordHandler for BofRecordHandler {
    /// Java `BofRecordHandler.processRecord` — parses type code; use [`Self::decide`].
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != BOF_SID || data.len() < 4 {
            return;
        }
        // bytes 2..4 = type (workbook=0x0005, worksheet=0x0010)
        let type_code = u16::from_le_bytes([data[2], data[3]]);
        let bof_type = match type_code {
            0x0005 => BofType::Workbook,
            0x0010 => BofType::Worksheet,
            _ => BofType::Other,
        };
        let _ = Self::decide(bof_type, None, false);
    }
}
