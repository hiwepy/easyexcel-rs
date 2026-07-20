//! Mirrors Java `com.alibaba.excel.analysis.v03.XlsSaxAnalyser`.

use std::path::{Path, PathBuf};

use easyexcel_core::{AnalysisContext, ExcelError, ExcelRow, ReadListener, Result};

use crate::analysis::excel_read_executor::ExcelReadExecutor;
use crate::context::{DefaultXlsReadContext, ReadSheet, XlsReadContext};
use crate::{list_xls_sheets, read_xls, ReadOptions};

/// Mirrors Java `XlsSaxAnalyser implements HSSFListener, ExcelReadExecutor`.
///
/// Java registers 19 BIFF record handlers and drives POI `HSSFEventFactory`.
/// Rust keeps the same public surface but delegates parsing to [`read_xls`] on
/// the calamine path.
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
    pub fn new(xls_read_context: DefaultXlsReadContext, path: impl Into<PathBuf>, options: ReadOptions) -> Result<Self> {
        let path = path.into();
        let discovered = list_xls_sheets(&path, &options)?;
        if discovered.is_empty() {
            return Err(ExcelError::Format("Can not find any sheet!".to_owned()));
        }
        let sheet_list = discovered
            .into_iter()
            .map(|(sheet_no, sheet_name)| ReadSheet::with_name(sheet_no, sheet_name))
            .collect();
        Ok(Self {
            path,
            options,
            xls_read_context,
            sheet_list,
            last_error: None,
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
    /// Returns `ExcelError::Unsupported` because BIFF dispatch is handled
    /// inside [`read_xls`] via calamine rather than POI record handlers.
    pub fn process_record(&self, _record_sid: u16, _data: &[u8]) -> Result<()> {
        Err(ExcelError::Unsupported(
            "XlsSaxAnalyser.processRecord is internal to read_xls calamine dispatch".to_owned(),
        ))
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
        let result = read_xls::<T, L>(&self.path, &self.options, listener);
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

    /// Mirrors Java `execute()`.
    ///
    /// Java pulls listeners from `ReadWorkbook`. Rust requires an explicit
    /// listener via [`execute_with_listener`](Self::execute_with_listener).
    fn execute(&mut self) {
        self.last_error = Some(ExcelError::Unsupported(
            "use XlsSaxAnalyser::execute_with_listener::<T, L>(...) to run read_xls".to_owned(),
        ));
    }
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::io::Read;

    use base64::Engine;
    use easyexcel_core::DynamicRow;
    use flate2::read::GzDecoder;
    use tempfile::{tempdir, NamedTempFile};

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
    fn void_execute_records_error_instead_of_panicking() {
        let file = write_java_multisheet_xls();
        let mut analyser =
            XlsSaxAnalyser::from_path(file.path(), ReadOptions::default()).expect("analyser");
        analyser.execute();
        assert!(analyser.last_error().is_some());
    }

    #[test]
    fn process_record_is_unsupported() {
        let file = write_java_multisheet_xls();
        let analyser =
            XlsSaxAnalyser::from_path(file.path(), ReadOptions::default()).expect("analyser");
        assert!(analyser.process_record(0x0203, &[]).is_err());
    }

    #[test]
    fn list_xls_sheets_matches_analyser_sheet_list() -> Result<()> {
        let file = write_java_multisheet_xls();
        let options = ReadOptions::default();
        let discovered = list_xls_sheets(file.path(), &options)?;
        let analyser = XlsSaxAnalyser::from_path(file.path(), options)?;
        assert_eq!(discovered.len(), analyser.sheet_list().len());
        Ok(())
    }

    #[test]
    fn empty_workbook_path_fails_sheet_discovery() {
        let directory = tempdir().expect("tempdir");
        let path = directory.path().join("missing.xls");
        assert!(list_xls_sheets(&path, &ReadOptions::default()).is_err());
    }
}
