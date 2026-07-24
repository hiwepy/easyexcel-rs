//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsSaxAnalyser`.

use std::path::{Path, PathBuf};

use easyexcel_core::{AnalysisContext, ExcelError, ExcelRow, ReadListener, Result};

use crate::analysis::excel_read_executor::ExcelReadExecutor;
use crate::analysis::v03::biff_record_stream::{read_workbook_stream, walk_biff_records};
use crate::analysis::v03::xls_list_sheet_listener::XlsListSheetListener;
use crate::analysis::v03::xls_record_dispatcher::{XlsRecordDispatchState, XlsRecordDispatcher};
use crate::context::{DefaultXlsReadContext, ReadSheet, XlsReadContext};
use crate::{ReadOptions, read_xls};

/// Mirrors Java `XlsSaxAnalyser implements HSSFListener, ExcelReadExecutor`.
///
/// Java registers 19 BIFF record handlers and drives POI `HSSFEventFactory`.
/// Rust runs the same SID dispatch over the real OLE Workbook stream, then uses
/// [`read_xls`] (calamine) for typed row materialisation.
pub struct XlsSaxAnalyser {
    /// Workbook path. (Java `ReadWorkbookHolder.file`)
    path: PathBuf,
    /// Read options collapsed from Java holders.
    options: ReadOptions,
    /// XLS read context. (Java `xlsReadContext`)
    xls_read_context: DefaultXlsReadContext,
    /// Discovered worksheets. (Java `sheetList`)
    sheet_list: Vec<ReadSheet>,
    /// Captures errors from the void [`ExcelReadExecutor::execute`] entry.
    last_error: Option<ExcelError>,
    /// Java-compatible BIFF SID handler registry and observable state.
    record_dispatcher: XlsRecordDispatcher,
}

impl XlsSaxAnalyser {
    /// Mirrors Java `XlsSaxAnalyser(XlsReadContext)`.
    ///
    /// Sheet discovery uses calamine metadata, equivalent to Java
    /// `XlsListSheetListener.execute()` pre-scan.
    ///
    /// # Errors
    ///
    /// Returns when the workbook cannot be opened or contains no sheets.
    pub fn new(
        mut xls_read_context: DefaultXlsReadContext,
        path: impl Into<PathBuf>,
        options: ReadOptions,
    ) -> Result<Self> {
        let path = path.into();
        let sheet_list = {
            let mut listener =
                XlsListSheetListener::new(&mut xls_read_context, &path, options.clone());
            listener.execute()?.to_vec()
        };
        if sheet_list.is_empty() {
            return Err(ExcelError::Format("Can not find any sheet!".to_owned()));
        }
        let record_dispatcher = XlsRecordDispatcher::new(&options);
        Ok(Self {
            path,
            options,
            xls_read_context,
            sheet_list,
            last_error: None,
            record_dispatcher,
        })
    }

    /// Convenience constructor mirroring Java `ExcelAnalyserImpl` wiring.
    ///
    /// # Errors
    ///
    /// Propagates [`Self::new`] failures.
    pub fn from_path(path: impl Into<PathBuf>, options: ReadOptions) -> Result<Self> {
        let context = DefaultXlsReadContext::new(&options);
        Self::new(context, path, options)
    }

    /// Returns the bound workbook path.
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Returns the XLS read context. (Java `xlsReadContext` field)
    #[must_use]
    pub fn xls_read_context(&self) -> &DefaultXlsReadContext {
        &self.xls_read_context
    }

    /// Returns the last error recorded by the void [`ExcelReadExecutor::execute`] entry.
    #[must_use]
    pub const fn last_error(&self) -> Option<&ExcelError> {
        self.last_error.as_ref()
    }

    /// Mirrors Java `processRecord(Record)` from `HSSFListener`.
    ///
    /// # Errors
    ///
    /// Routes the record to the registered handler. Unknown SIDs are ignored,
    /// matching Java's `XLS_RECORD_HANDLER_MAP.get(...) == null` branch.
    pub fn process_record(&mut self, record_sid: u16, data: &[u8]) -> Result<()> {
        self.record_dispatcher.process_record(record_sid, data)
    }

    /// Returns state produced by the real BIFF handler dispatch.
    #[must_use]
    pub const fn record_dispatch_state(&self) -> &XlsRecordDispatchState {
        self.record_dispatcher.state()
    }

    fn dispatch_workbook_records(&mut self) -> Result<()> {
        let workbook = read_workbook_stream(&self.path)?;
        self.record_dispatcher.reset();
        walk_biff_records(&workbook, |record_sid, data| {
            self.process_record(record_sid, data)
        })?;
        self.record_dispatcher.finish_records()
    }

    /// Typed execute path. (Java `execute()` + listener on `ReadWorkbook`)
    ///
    /// # Errors
    ///
    /// Propagates workbook, sheet-selection, conversion, or listener errors.
    pub fn execute_with_listener<T, L>(&mut self, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        let options = self.options.clone();
        self.execute::<T, L>(&options, listener)
    }

    fn execute_with_options<T, L>(&mut self, options: &ReadOptions, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        let result = self
            .dispatch_workbook_records()
            .and_then(|()| read_xls::<T, L>(&self.path, options, listener));
        match result {
            Ok(()) => {
                self.last_error = None;
                Ok(())
            }
            Err(error) => {
                self.last_error = Some(error.clone());
                Err(error)
            }
        }
    }

    /// Returns the listener callback context from the embedded read context.
    #[must_use]
    pub fn analysis_context(&self) -> &AnalysisContext {
        self.xls_read_context
            .analysis_context_impl()
            .analysis_context()
    }
}

impl ExcelReadExecutor for XlsSaxAnalyser {
    /// Mirrors Java `sheetList()`.
    fn sheet_list(&self) -> &[ReadSheet] {
        &self.sheet_list
    }

    fn execute<T, L>(&mut self, options: &ReadOptions, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        self.execute_with_options::<T, L>(options, listener)
    }
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::Read;

    use base64::Engine;
    use easyexcel_core::DynamicRow;
    use flate2::read::GzDecoder;
    use tempfile::{NamedTempFile, tempdir};

    use super::*;

    #[derive(Default)]
    struct CollectingListener {
        rows: Vec<DynamicRow>,
    }

    impl ReadListener<DynamicRow> for CollectingListener {
        fn invoke(&mut self, data: DynamicRow, _context: &AnalysisContext) -> Result<()> {
            self.rows.push(data);
            Ok(())
        }

        fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
            Ok(())
        }
    }

    /// Materializes the embedded Java multisheet `.xls` fixture for unit tests.
    fn write_java_multisheet_xls() -> NamedTempFile {
        let file = NamedTempFile::with_suffix(".xls").expect("temp xls");
        let compressed = base64::engine::general_purpose::STANDARD
            .decode(include_str!("../../fixtures/java-multiplesheets.xls.gz.b64").trim())
            .expect("fixture b64");
        let mut decoder = GzDecoder::new(compressed.as_slice());
        let mut workbook = Vec::new();
        decoder.read_to_end(&mut workbook).expect("gunzip");
        fs::write(file.path(), workbook).expect("write xls");
        file
    }

    #[test]
    fn sheet_list_discovers_worksheet_names() -> Result<()> {
        let file = write_java_multisheet_xls();
        let options = ReadOptions::default();
        let analyser = XlsSaxAnalyser::from_path(file.path(), options)?;
        assert_eq!(analyser.sheet_list().len(), 6);
        assert!(!analyser.sheet_list()[0].sheet_name().is_empty());
        assert!(
            !analyser
                .xls_read_context()
                .xls_read_workbook_holder()
                .need_read_sheet()
        );
        assert_eq!(
            analyser
                .xls_read_context()
                .xls_read_workbook_holder()
                .inner()
                .actual_sheet_data_list()
                .expect("actual sheet list"),
            analyser.sheet_list()
        );
        Ok(())
    }

    #[test]
    fn execute_with_listener_delegates_to_read_xls() -> Result<()> {
        let file = write_java_multisheet_xls();
        let mut options = ReadOptions::default();
        options.head_row_number = 1;
        options.sheet = crate::SheetSelector::Index(0);
        let mut analyser = XlsSaxAnalyser::from_path(file.path(), options)?;
        let mut listener = CollectingListener::default();
        analyser.execute_with_listener::<DynamicRow, _>(&mut listener)?;
        assert!(!listener.rows.is_empty());
        Ok(())
    }

    #[test]
    fn trait_execute_runs_the_real_xls_parser() -> Result<()> {
        let file = write_java_multisheet_xls();
        let mut options = ReadOptions::default();
        options.head_row_number = 1;
        options.sheet = crate::SheetSelector::Index(0);
        let mut analyser =
            XlsSaxAnalyser::from_path(file.path(), options.clone()).expect("analyser");
        let mut listener = CollectingListener::default();
        ExcelReadExecutor::execute::<DynamicRow, _>(&mut analyser, &options, &mut listener)?;
        assert!(!listener.rows.is_empty());
        assert!(analyser.last_error().is_none());
        Ok(())
    }

    #[test]
    fn process_record_dispatches_number_handler() -> Result<()> {
        let file = write_java_multisheet_xls();
        let mut analyser =
            XlsSaxAnalyser::from_path(file.path(), ReadOptions::default()).expect("analyser");
        let mut payload = vec![4, 0, 2, 0, 9, 0];
        payload.extend_from_slice(&12.5f64.to_le_bytes());
        analyser.process_record(0x0203, &payload)?;
        let cell = analyser
            .record_dispatch_state()
            .last_number_cell()
            .expect("real NumberRecordHandler output");
        assert_eq!((cell.row, cell.column, cell.format_index), (4, 2, 9));
        assert_eq!(cell.value, 12.5);
        Ok(())
    }

    #[test]
    fn execute_walks_real_workbook_biff_records() -> Result<()> {
        let file = write_java_multisheet_xls();
        let mut options = ReadOptions::default();
        options.head_row_number = 1;
        options.sheet = crate::SheetSelector::Index(0);
        let mut analyser = XlsSaxAnalyser::from_path(file.path(), options)?;
        let mut listener = CollectingListener::default();

        analyser.execute_with_listener::<DynamicRow, _>(&mut listener)?;

        let state = analyser.record_dispatch_state();
        assert!(state.total_record_count() > 0);
        assert!(state.handled_record_count() > 0);
        assert_eq!(state.bound_sheets().len(), analyser.sheet_list().len());
        assert_eq!(state.worksheet_bof_count(), analyser.sheet_list().len());
        assert!(state.workbook_bof_count() >= 1);
        assert!(state.eof_count() >= analyser.sheet_list().len());
        assert!(!state.shared_strings().is_empty());
        assert_eq!(
            state.unique_string_count(),
            u32::try_from(state.shared_strings().len()).ok()
        );
        assert!(
            state
                .bound_sheets()
                .iter()
                .all(|sheet| !sheet.name.is_empty())
        );
        Ok(())
    }

    #[test]
    fn list_xls_sheets_matches_analyser_sheet_list() -> Result<()> {
        let file = write_java_multisheet_xls();
        let options = ReadOptions::default();
        let discovered = crate::list_xls_sheets(file.path(), &options)?;
        let analyser = XlsSaxAnalyser::from_path(file.path(), options)?;
        assert_eq!(discovered.len(), analyser.sheet_list().len());
        Ok(())
    }

    #[test]
    fn empty_workbook_path_fails_sheet_discovery() {
        let directory = tempdir().expect("tempdir");
        let path = directory.path().join("missing.xls");
        assert!(crate::list_xls_sheets(&path, &ReadOptions::default()).is_err());
    }
}
