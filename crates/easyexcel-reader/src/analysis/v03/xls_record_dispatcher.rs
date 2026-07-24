//! BIFF SID-to-handler dispatch matching Java `XlsSaxAnalyser.processRecord`.

use easyexcel_core::{CellExtraType, Result};

use crate::{ReadOptions, SheetSelector};

use super::biff_string::decode_sst_segments;
use super::handlers::blank_record_handler::{BLANK_SID, BlankCell, BlankRecordHandler};
use super::handlers::bof_record_handler::{BOF_SID, BofRecordHandler};
use super::handlers::bool_err_record_handler::{BOOL_ERR_SID, BoolCell, BoolErrRecordHandler};
use super::handlers::bound_sheet_record_handler::{
    BOUND_SHEET_SID, BoundSheetEntry, BoundSheetRecordHandler,
};
use super::handlers::dummy_record_handler::DummyRecordHandler;
use super::handlers::eof_record_handler::{EOF_SID, EofRecordHandler};
use super::handlers::formula_record_handler::{FORMULA_SID, FormulaCell, FormulaRecordHandler};
use super::handlers::hyperlink_record_handler::HyperlinkRecordHandler;
use super::handlers::index_record_handler::{INDEX_SID, IndexRecordHandler};
use super::handlers::label_record_handler::{LABEL_SID, LabelRecordHandler};
use super::handlers::label_sst_record_handler::{
    LABEL_SST_SID, LabelSstCell, LabelSstRecordHandler,
};
use super::handlers::merge_cells_record_handler::MergeCellsRecordHandler;
use super::handlers::note_record_handler::NoteRecordHandler;
use super::handlers::number_record_handler::{NUMBER_SID, NumberCell, NumberRecordHandler};
use super::handlers::obj_record_handler::{OBJ_SID, ObjRecordHandler};
use super::handlers::rk_record_handler::{RK_SID, RkRecordHandler};
use super::handlers::sst_record_handler::{SST_SID, SstRecordHandler};
use super::handlers::string_record_handler::{STRING_SID, StringRecordHandler};
use super::handlers::text_object_record_handler::{TEXT_OBJECT_SID, TextObjectRecordHandler};
use super::xls_record_handler::XlsRecordHandler;

const HYPERLINK_SID: u16 = 0x01B8;
const MERGE_CELLS_SID: u16 = 0x00E5;
const NOTE_SID: u16 = 0x001C;
const DUMMY_RECORD_SID: u16 = u16::MAX;
const CONTINUE_SID: u16 = 0x003C;

/// Observable result of running Java-compatible BIFF handler dispatch.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct XlsRecordDispatchState {
    total_record_count: usize,
    handled_record_count: usize,
    unknown_record_count: usize,
    skipped_record_count: usize,
    workbook_bof_count: usize,
    worksheet_bof_count: usize,
    eof_count: usize,
    bound_sheets: Vec<BoundSheetEntry>,
    unique_string_count: Option<u32>,
    approximate_total_row_number: Option<u32>,
    last_blank_cell: Option<BlankCell>,
    last_boolean_cell: Option<BoolCell>,
    last_number_cell: Option<NumberCell>,
    last_rk_cell: Option<BlankCell>,
    shared_strings: Vec<String>,
    last_label_sst_cell: Option<LabelSstCell>,
    last_formula_cell: Option<FormulaCell>,
}

impl XlsRecordDispatchState {
    /// Number of physical BIFF records presented to the dispatcher.
    #[must_use]
    pub const fn total_record_count(&self) -> usize {
        self.total_record_count
    }

    /// Number of records routed to a registered handler.
    #[must_use]
    pub const fn handled_record_count(&self) -> usize {
        self.handled_record_count
    }

    /// Number of records ignored because Java has no registered handler SID.
    #[must_use]
    pub const fn unknown_record_count(&self) -> usize {
        self.unknown_record_count
    }

    /// Number of known records skipped by a disabled `support()` predicate.
    #[must_use]
    pub const fn skipped_record_count(&self) -> usize {
        self.skipped_record_count
    }

    /// Number of workbook-global BOF records.
    #[must_use]
    pub const fn workbook_bof_count(&self) -> usize {
        self.workbook_bof_count
    }

    /// Number of worksheet BOF records.
    #[must_use]
    pub const fn worksheet_bof_count(&self) -> usize {
        self.worksheet_bof_count
    }

    /// Number of EOF records.
    #[must_use]
    pub const fn eof_count(&self) -> usize {
        self.eof_count
    }

    /// Bound-sheet metadata decoded by `BoundSheetRecordHandler`.
    #[must_use]
    pub fn bound_sheets(&self) -> &[BoundSheetEntry] {
        &self.bound_sheets
    }

    /// Unique string count announced by the SST record.
    #[must_use]
    pub const fn unique_string_count(&self) -> Option<u32> {
        self.unique_string_count
    }

    /// Approximate row count announced by the last Index record.
    #[must_use]
    pub const fn approximate_total_row_number(&self) -> Option<u32> {
        self.approximate_total_row_number
    }

    /// Last blank cell decoded by the registered handler.
    #[must_use]
    pub const fn last_blank_cell(&self) -> Option<BlankCell> {
        self.last_blank_cell
    }

    /// Last boolean cell decoded by the registered handler.
    #[must_use]
    pub const fn last_boolean_cell(&self) -> Option<BoolCell> {
        self.last_boolean_cell
    }

    /// Last number cell decoded by the registered handler.
    #[must_use]
    pub fn last_number_cell(&self) -> Option<&NumberCell> {
        self.last_number_cell.as_ref()
    }

    /// Last RK placement decoded with EasyExcel's historical empty-cell rule.
    #[must_use]
    pub const fn last_rk_cell(&self) -> Option<BlankCell> {
        self.last_rk_cell
    }

    /// Fully decoded shared-string table in BIFF index order.
    #[must_use]
    pub fn shared_strings(&self) -> &[String] {
        &self.shared_strings
    }

    /// Last LabelSST cell resolved through the decoded shared-string table.
    #[must_use]
    pub const fn last_label_sst_cell(&self) -> Option<&LabelSstCell> {
        self.last_label_sst_cell.as_ref()
    }

    /// Last completed cached formula result.
    #[must_use]
    pub const fn last_formula_cell(&self) -> Option<&FormulaCell> {
        self.last_formula_cell.as_ref()
    }
}

/// Owns the 19 Java EasyExcel XLS handlers and dispatches by BIFF SID.
#[derive(Debug)]
pub struct XlsRecordDispatcher {
    state: XlsRecordDispatchState,
    blank: BlankRecordHandler,
    bof: BofRecordHandler,
    bool_err: BoolErrRecordHandler,
    bound_sheet: BoundSheetRecordHandler,
    dummy: DummyRecordHandler,
    eof: EofRecordHandler,
    formula: FormulaRecordHandler,
    hyperlink: HyperlinkRecordHandler,
    index: IndexRecordHandler,
    label: LabelRecordHandler,
    label_sst: LabelSstRecordHandler,
    merge_cells: MergeCellsRecordHandler,
    note: NoteRecordHandler,
    number: NumberRecordHandler,
    obj: ObjRecordHandler,
    rk: RkRecordHandler,
    sst: SstRecordHandler,
    string: StringRecordHandler,
    text_object: TextObjectRecordHandler,
    sheet_selector: SheetSelector,
    next_sheet_index: usize,
    ignore_record: bool,
    auto_trim: bool,
    pending_sst_segments: Option<Vec<Vec<u8>>>,
    pending_formula_string_segments: Option<Vec<Vec<u8>>>,
}

impl XlsRecordDispatcher {
    /// Creates the handler map using Java `support(context)` feature flags.
    #[must_use]
    pub fn new(options: &ReadOptions) -> Self {
        Self {
            state: XlsRecordDispatchState::default(),
            blank: BlankRecordHandler::new(),
            bof: BofRecordHandler::new(),
            bool_err: BoolErrRecordHandler::new(),
            bound_sheet: BoundSheetRecordHandler::new(),
            dummy: DummyRecordHandler::new(),
            eof: EofRecordHandler::new(),
            formula: FormulaRecordHandler::new(),
            hyperlink: HyperlinkRecordHandler::new(
                options.extra_read.contains(&CellExtraType::Hyperlink),
            ),
            index: IndexRecordHandler::new(),
            label: LabelRecordHandler::new(),
            label_sst: LabelSstRecordHandler::new(),
            merge_cells: MergeCellsRecordHandler::new(
                options.extra_read.contains(&CellExtraType::Merge),
            ),
            note: NoteRecordHandler::new(options.extra_read.contains(&CellExtraType::Comment)),
            number: NumberRecordHandler::new(),
            obj: ObjRecordHandler::new(),
            rk: RkRecordHandler::new(),
            sst: SstRecordHandler::new(),
            string: StringRecordHandler::new(),
            text_object: TextObjectRecordHandler::new(),
            sheet_selector: options.sheet.clone(),
            next_sheet_index: 0,
            ignore_record: false,
            auto_trim: options.auto_trim,
            pending_sst_segments: None,
            pending_formula_string_segments: None,
        }
    }

    /// Resets all per-workbook state while preserving configured feature flags.
    pub fn reset(&mut self) {
        let hyperlink_enabled = self.hyperlink.enabled;
        let merge_enabled = self.merge_cells.enabled;
        let note_enabled = self.note.enabled;
        let sheet_selector = self.sheet_selector.clone();
        let auto_trim = self.auto_trim;
        *self = Self {
            state: XlsRecordDispatchState::default(),
            blank: BlankRecordHandler::new(),
            bof: BofRecordHandler::new(),
            bool_err: BoolErrRecordHandler::new(),
            bound_sheet: BoundSheetRecordHandler::new(),
            dummy: DummyRecordHandler::new(),
            eof: EofRecordHandler::new(),
            formula: FormulaRecordHandler::new(),
            hyperlink: HyperlinkRecordHandler::new(hyperlink_enabled),
            index: IndexRecordHandler::new(),
            label: LabelRecordHandler::new(),
            label_sst: LabelSstRecordHandler::new(),
            merge_cells: MergeCellsRecordHandler::new(merge_enabled),
            note: NoteRecordHandler::new(note_enabled),
            number: NumberRecordHandler::new(),
            obj: ObjRecordHandler::new(),
            rk: RkRecordHandler::new(),
            sst: SstRecordHandler::new(),
            string: StringRecordHandler::new(),
            text_object: TextObjectRecordHandler::new(),
            sheet_selector,
            next_sheet_index: 0,
            ignore_record: false,
            auto_trim,
            pending_sst_segments: None,
            pending_formula_string_segments: None,
        };
    }

    /// Returns observable dispatch state for diagnostics and parity tests.
    #[must_use]
    pub const fn state(&self) -> &XlsRecordDispatchState {
        &self.state
    }

    /// Mirrors Java `XlsSaxAnalyser.processRecord`: unknown SIDs are ignored,
    /// disabled handlers are skipped, and known records reach their handler.
    pub fn process_record(&mut self, record_sid: u16, data: &[u8]) -> Result<()> {
        self.state.total_record_count += 1;
        if record_sid == CONTINUE_SID {
            if let Some(segments) = self.pending_sst_segments.as_mut() {
                segments.push(data.to_vec());
                self.try_finalize_sst(false)?;
                return Ok(());
            }
            if let Some(segments) = self.pending_formula_string_segments.as_mut() {
                segments.push(data.to_vec());
                self.try_finalize_formula_string(false)?;
                return Ok(());
            }
            self.state.unknown_record_count += 1;
            return Ok(());
        }
        self.finish_pending_records()?;
        if self.ignore_record && is_ignorable_sid(record_sid) {
            self.state.skipped_record_count += 1;
            return Ok(());
        }
        match record_sid {
            BLANK_SID => self.dispatch_blank(record_sid, data),
            BOF_SID => {
                if data.len() >= 4 {
                    match u16::from_le_bytes([data[2], data[3]]) {
                        0x0005 => {
                            self.state.workbook_bof_count += 1;
                            self.next_sheet_index = 0;
                            self.ignore_record = false;
                        }
                        0x0010 => {
                            self.state.worksheet_bof_count += 1;
                            self.ignore_record = !self.should_read_sheet(self.next_sheet_index);
                            self.next_sheet_index = self.next_sheet_index.saturating_add(1);
                        }
                        _ => {}
                    }
                }
                self.bof.process_record(record_sid, data);
            }
            BOOL_ERR_SID => self.dispatch_bool(record_sid, data),
            BOUND_SHEET_SID => {
                self.bound_sheet.process_record(record_sid, data);
                self.state.bound_sheets = self.bound_sheet.ordered_sheets();
            }
            DUMMY_RECORD_SID => self.dummy.process_record(record_sid, data),
            EOF_SID => {
                self.state.eof_count += 1;
                self.eof.process_record(record_sid, data);
            }
            FORMULA_SID => {
                self.formula.process_record(record_sid, data);
                self.state.last_formula_cell = self.formula.last_cell.clone();
            }
            HYPERLINK_SID => {
                if !self.hyperlink.support() {
                    self.state.skipped_record_count += 1;
                    return Ok(());
                }
                self.hyperlink.process_record(record_sid, data);
            }
            INDEX_SID => {
                self.index.process_record(record_sid, data);
                self.state.approximate_total_row_number = self.index.approximate_total_row_number;
            }
            LABEL_SID => self.label.process_record(record_sid, data),
            LABEL_SST_SID => {
                self.label_sst.process_record(record_sid, data);
                if let Some(reference) = self.label_sst.last_reference {
                    let cell = LabelSstRecordHandler::process_label_sst(
                        reference.row,
                        reference.column,
                        reference.sst_index,
                        self.auto_trim,
                        &|index| self.sst.get(index).map(str::to_owned),
                    );
                    self.state.last_label_sst_cell = Some(cell);
                }
            }
            MERGE_CELLS_SID => {
                if !self.merge_cells.support() {
                    self.state.skipped_record_count += 1;
                    return Ok(());
                }
                self.merge_cells.process_record(record_sid, data);
            }
            NOTE_SID => {
                if !self.note.support() {
                    self.state.skipped_record_count += 1;
                    return Ok(());
                }
                self.note.process_record(record_sid, data);
            }
            NUMBER_SID => self.dispatch_number(record_sid, data),
            OBJ_SID => self.obj.process_record(record_sid, data),
            RK_SID => self.dispatch_rk(record_sid, data),
            SST_SID => {
                self.sst.process_record(record_sid, data);
                self.state.unique_string_count = self.sst.unique_string_count;
                self.pending_sst_segments = Some(vec![data.to_vec()]);
                self.try_finalize_sst(false)?;
            }
            STRING_SID => {
                self.pending_formula_string_segments = Some(vec![data.to_vec()]);
                self.try_finalize_formula_string(false)?;
            }
            TEXT_OBJECT_SID => self.text_object.process_record(record_sid, data),
            _ => {
                self.state.unknown_record_count += 1;
                return Ok(());
            }
        }
        self.state.handled_record_count += 1;
        Ok(())
    }

    /// Finalizes a continuable logical record at end-of-stream.
    pub fn finish_records(&mut self) -> Result<()> {
        self.finish_pending_records()
    }

    fn dispatch_blank(&mut self, record_sid: u16, data: &[u8]) {
        self.blank.process_record(record_sid, data);
        self.state.last_blank_cell = self.blank.last_cell;
    }

    fn dispatch_bool(&mut self, record_sid: u16, data: &[u8]) {
        self.bool_err.process_record(record_sid, data);
        self.state.last_boolean_cell = self.bool_err.last_cell;
    }

    fn dispatch_number(&mut self, record_sid: u16, data: &[u8]) {
        self.number.process_record(record_sid, data);
        self.state.last_number_cell = self.number.last_cell.clone();
    }

    fn dispatch_rk(&mut self, record_sid: u16, data: &[u8]) {
        self.rk.process_record(record_sid, data);
        self.state.last_rk_cell = self.rk.last_cell;
    }

    fn should_read_sheet(&self, index: usize) -> bool {
        match &self.sheet_selector {
            SheetSelector::First => index == 0,
            SheetSelector::Index(selected) => index == *selected,
            SheetSelector::Name(selected) => self
                .state
                .bound_sheets
                .get(index)
                .is_some_and(|sheet| sheet.name == *selected),
            SheetSelector::All => true,
        }
    }

    fn finish_pending_records(&mut self) -> Result<()> {
        self.try_finalize_sst(true)?;
        self.try_finalize_formula_string(true)
    }

    fn try_finalize_sst(&mut self, require_complete: bool) -> Result<()> {
        let Some(segments) = self.pending_sst_segments.as_ref() else {
            return Ok(());
        };
        match decode_sst_segments(segments) {
            Ok(strings) => {
                let unique = u32::try_from(strings.len()).map_err(|_| {
                    easyexcel_core::ExcelError::Format(
                        "decoded SST size exceeds BIFF u32 range".to_owned(),
                    )
                })?;
                self.pending_sst_segments = None;
                self.sst.process_decoded_sst(unique, strings.clone());
                self.state.unique_string_count = Some(unique);
                self.state.shared_strings = strings;
                Ok(())
            }
            Err(_) if !require_complete => Ok(()),
            Err(error) => Err(error),
        }
    }

    fn try_finalize_formula_string(&mut self, require_complete: bool) -> Result<()> {
        let Some(segments) = self.pending_formula_string_segments.as_ref() else {
            return Ok(());
        };
        match StringRecordHandler::decode_segments(segments) {
            Ok(value) => {
                self.pending_formula_string_segments = None;
                self.string.process_decoded(value.clone());
                if let Some((cell, _)) =
                    StringRecordHandler::process_string(&mut self.formula, value, self.auto_trim)
                {
                    self.state.last_formula_cell = Some(cell);
                }
                Ok(())
            }
            Err(_) if !require_complete => Ok(()),
            Err(error) => Err(error),
        }
    }
}

fn is_ignorable_sid(record_sid: u16) -> bool {
    matches!(
        record_sid,
        BLANK_SID
            | BOOL_ERR_SID
            | BOUND_SHEET_SID
            | DUMMY_RECORD_SID
            | FORMULA_SID
            | HYPERLINK_SID
            | INDEX_SID
            | LABEL_SID
            | LABEL_SST_SID
            | MERGE_CELLS_SID
            | NOTE_SID
            | NUMBER_SID
            | OBJ_SID
            | RK_SID
            | SST_SID
            | STRING_SID
            | TEXT_OBJECT_SID
    )
}

impl Default for XlsRecordDispatcher {
    fn default() -> Self {
        Self::new(&ReadOptions::default())
    }
}

#[cfg(test)]
mod tests {
    use super::super::handlers::formula_record_handler::FormulaCachedType;
    use super::*;

    #[test]
    fn dispatches_number_to_real_handler_and_keeps_unknown_records_ignorable() -> Result<()> {
        let mut dispatcher = XlsRecordDispatcher::default();
        let mut number = vec![2, 0, 3, 0, 7, 0];
        number.extend_from_slice(&42.5f64.to_le_bytes());

        dispatcher.process_record(NUMBER_SID, &number)?;
        dispatcher.process_record(0x1234, &[])?;

        assert_eq!(dispatcher.state().total_record_count(), 2);
        assert_eq!(dispatcher.state().handled_record_count(), 1);
        assert_eq!(dispatcher.state().unknown_record_count(), 1);
        let cell = dispatcher
            .state()
            .last_number_cell()
            .expect("number handler output");
        assert_eq!((cell.row, cell.column, cell.format_index), (2, 3, 7));
        assert_eq!(cell.value, 42.5);
        Ok(())
    }

    #[test]
    fn support_predicate_skips_merge_when_not_requested() -> Result<()> {
        let mut dispatcher = XlsRecordDispatcher::default();
        dispatcher.process_record(MERGE_CELLS_SID, &[0; 10])?;
        assert_eq!(dispatcher.state().handled_record_count(), 0);
        assert_eq!(dispatcher.state().skipped_record_count(), 1);
        Ok(())
    }

    #[test]
    fn selected_sheet_skips_ignorable_records_until_next_bof() -> Result<()> {
        let mut dispatcher = XlsRecordDispatcher::default();
        let workbook_bof = [0, 0, 0x05, 0x00];
        let worksheet_bof = [0, 0, 0x10, 0x00];
        let mut first_number = vec![0, 0, 0, 0, 0, 0];
        first_number.extend_from_slice(&1.0f64.to_le_bytes());
        let mut second_number = vec![0, 0, 0, 0, 0, 0];
        second_number.extend_from_slice(&2.0f64.to_le_bytes());

        dispatcher.process_record(BOF_SID, &workbook_bof)?;
        dispatcher.process_record(BOF_SID, &worksheet_bof)?;
        dispatcher.process_record(NUMBER_SID, &first_number)?;
        dispatcher.process_record(EOF_SID, &[])?;
        dispatcher.process_record(BOF_SID, &worksheet_bof)?;
        dispatcher.process_record(NUMBER_SID, &second_number)?;

        assert_eq!(
            dispatcher
                .state()
                .last_number_cell()
                .expect("first sheet number")
                .value,
            1.0
        );
        assert_eq!(dispatcher.state().skipped_record_count(), 1);
        Ok(())
    }

    #[test]
    fn sst_continue_resolves_following_label_sst() -> Result<()> {
        let mut dispatcher = XlsRecordDispatcher::default();
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&4u16.to_le_bytes());
        sst.push(0);
        sst.extend_from_slice(b"  ");
        dispatcher.process_record(SST_SID, &sst)?;
        dispatcher.process_record(CONTINUE_SID, &[0, b'o', b'k'])?;

        let mut label = vec![3, 0, 2, 0, 0, 0];
        label.extend_from_slice(&0u32.to_le_bytes());
        dispatcher.process_record(LABEL_SST_SID, &label)?;

        assert_eq!(dispatcher.state().shared_strings(), &["  ok".to_owned()]);
        assert_eq!(
            dispatcher.state().last_label_sst_cell(),
            Some(&LabelSstCell::String {
                row: 3,
                column: 2,
                value: "ok".to_owned(),
            })
        );
        Ok(())
    }

    #[test]
    fn formula_string_record_completes_pending_cached_result_across_continue() -> Result<()> {
        let mut dispatcher = XlsRecordDispatcher::default();
        let formula = vec![5, 0, 4, 0, 0, 0, 0x00, 0, 0, 0, 0, 0, 0xFF, 0xFF];
        dispatcher.process_record(FORMULA_SID, &formula)?;
        assert!(
            dispatcher
                .state()
                .last_formula_cell()
                .is_none_or(|cell| cell.cached_type != FormulaCachedType::String)
        );

        let string = vec![4, 0, 0, b'a', b'b'];
        dispatcher.process_record(STRING_SID, &string)?;
        dispatcher.process_record(CONTINUE_SID, &[0, b'c', b'd'])?;

        let cell = dispatcher
            .state()
            .last_formula_cell()
            .expect("completed string formula");
        assert_eq!((cell.row, cell.column), (5, 4));
        assert_eq!(cell.string_value.as_deref(), Some("abcd"));
        assert!(!cell.pending_string);
        Ok(())
    }

    #[test]
    fn finish_records_rejects_truncated_continuable_record() -> Result<()> {
        let mut dispatcher = XlsRecordDispatcher::default();
        let mut sst = Vec::new();
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&1u32.to_le_bytes());
        sst.extend_from_slice(&2u16.to_le_bytes());
        sst.push(0);
        sst.push(b'a');
        dispatcher.process_record(SST_SID, &sst)?;
        assert!(dispatcher.finish_records().is_err());
        Ok(())
    }
}
