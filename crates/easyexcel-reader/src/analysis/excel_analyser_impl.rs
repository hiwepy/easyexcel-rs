//! Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyserImpl`.

use std::path::{Path, PathBuf};
use std::sync::Arc;

use easyexcel_core::support::ExcelTypeEnum;
use easyexcel_core::{
    AnalysisContext, CellExtra, ErrorAction, ExcelError, ExcelRow, ReadListener, Result,
};

use crate::analysis::excel_analyser::ExcelAnalyser;
use crate::analysis::excel_read_executor::ExcelReadExecutorKind;
use crate::{ReadOptions, SheetSelector};

struct ContextTrackingReadListener<'a, L> {
    delegate: &'a mut L,
    latest: &'a mut AnalysisContext,
}

impl<L> ContextTrackingReadListener<'_, L> {
    fn capture(&mut self, context: &AnalysisContext) {
        self.latest.clone_from(context);
    }
}

impl<T, L> ReadListener<T> for ContextTrackingReadListener<'_, L>
where
    L: ReadListener<T>,
{
    fn on_exception(&mut self, error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        self.capture(context);
        self.delegate.on_exception(error, context)
    }

    fn invoke_head(
        &mut self,
        head: &std::collections::HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        self.capture(context);
        self.delegate.invoke_head(head, context)
    }

    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        self.capture(context);
        self.delegate.invoke(data, context)
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        self.capture(context);
        self.delegate.extra(extra, context)
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.capture(context);
        self.delegate.do_after_all_analysed(context)
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        self.capture(context);
        self.delegate.has_next(context)
    }
}

/// Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyserImpl implements ExcelAnalyser`.
///
/// Java constructs the analyser from `ReadWorkbook`, then
/// [`choiceExcelExecutor`](ExcelAnalyserImpl::choice_excel_executor) selects
/// `XlsxSaxAnalyser` / `XlsSaxAnalyser` / `CsvExcelReadExecutor` by
/// `ExcelTypeEnum`. Rust keeps the same dispatch table but delegates the
/// actual parse to [`read_xlsx`] / [`read_xls`] / [`read_csv`].
pub struct ExcelAnalyserImpl {
    /// Workbook path. (Java `ReadWorkbook.file`)
    path: Option<PathBuf>,
    /// Read options collapsed from Java `ReadWorkbook` + holders.
    options: ReadOptions,
    /// Detected workbook type. (Java `ExcelTypeEnum valueOf(ReadWorkbook)`)
    excel_type: Option<ExcelTypeEnum>,
    /// Format-specific executor selected during construction.
    excel_read_executor: Option<ExcelReadExecutorKind>,
    /// Live analysis context. (Java `ExcelAnalyserImpl.analysisContext`)
    context: AnalysisContext,
    /// Prevents multiple shutdowns. (Java `ExcelAnalyserImpl.finished`)
    finished: bool,
    /// Keeps a materialised Java-style input stream alive until `finish`.
    temporary_input: Option<Arc<tempfile::TempPath>>,
    /// Captures the last typed analysis error.
    last_error: Option<ExcelError>,
}

impl Default for ExcelAnalyserImpl {
    /// Creates an idle analyser with a safe empty context (no `unreachable!`).
    fn default() -> Self {
        Self::new()
    }
}

impl ExcelAnalyserImpl {
    /// Creates an idle analyser.
    ///
    /// Prefer [`from_path`](Self::from_path) for the Java
    /// `ExcelAnalyserImpl(ReadWorkbook)` hot path.
    #[must_use]
    pub fn new() -> Self {
        Self {
            path: None,
            options: ReadOptions::default(),
            excel_type: None,
            excel_read_executor: None,
            // Safe default — Java leaves `analysisContext` null until
            // `choiceExcelExecutor`; Rust always exposes a usable value.
            context: AnalysisContext::new("", 0, 0),
            finished: false,
            temporary_input: None,
            last_error: None,
        }
    }

    /// Java `ExcelAnalyserImpl(ReadWorkbook)` — binds a path and chooses the executor.
    ///
    /// # Errors
    ///
    /// Returns when the workbook type cannot be resolved from the path.
    pub fn from_path(path: impl Into<PathBuf>, options: ReadOptions) -> Result<Self> {
        let mut analyser = Self {
            path: Some(path.into()),
            options,
            excel_type: None,
            excel_read_executor: None,
            context: AnalysisContext::new("", 0, 0),
            finished: false,
            temporary_input: None,
            last_error: None,
        };
        analyser.choice_excel_executor()?;
        Ok(analyser)
    }

    /// Creates an analyser whose workbook path is owned by a temporary-file guard.
    ///
    /// Java accepts non-seekable `InputStream` values and deletes the
    /// materialised file from `finish()`. Keeping the guard in the analyser
    /// also guarantees cleanup when analysis itself fails and calls `finish`.
    pub(crate) fn from_temporary_input(
        path: impl Into<PathBuf>,
        temporary_input: Arc<tempfile::TempPath>,
        options: ReadOptions,
    ) -> Result<Self> {
        let mut analyser = Self::from_path(path, options)?;
        analyser.temporary_input = Some(temporary_input);
        Ok(analyser)
    }

    /// Returns the bound workbook path, if any.
    #[must_use]
    pub fn path(&self) -> Option<&Path> {
        self.path.as_deref()
    }

    /// Returns the resolved Excel type after [`choice_excel_executor`].
    #[must_use]
    pub const fn excel_type(&self) -> Option<ExcelTypeEnum> {
        self.excel_type
    }

    /// Returns the last error recorded by the void [`ExcelAnalyser::analysis`] entry.
    #[must_use]
    pub const fn last_error(&self) -> Option<&ExcelError> {
        self.last_error.as_ref()
    }

    /// Returns whether [`finish`](ExcelAnalyser::finish) has run.
    #[must_use]
    pub const fn is_finished(&self) -> bool {
        self.finished
    }

    /// Returns whether this analyser owns a materialised input stream.
    #[must_use]
    pub const fn has_temporary_input(&self) -> bool {
        self.temporary_input.is_some()
    }

    /// Returns the bound read options.
    #[must_use]
    pub const fn options(&self) -> &ReadOptions {
        &self.options
    }

    /// Returns mutable read options for sheet-scoped reads.
    pub fn options_mut(&mut self) -> &mut ReadOptions {
        &mut self.options
    }

    /// Updates the active sheet selector before analysis.
    pub fn set_sheet_selector(&mut self, sheet: SheetSelector) {
        self.options.sheet = sheet;
    }

    /// Mirrors Java `choiceExcelExecutor(ReadWorkbook)`.
    ///
    /// Resolves `ExcelTypeEnum` from the path extension (with a CSV fallback)
    /// and seeds [`analysis_context`](ExcelAnalyser::analysis_context). It
    /// constructs the concrete XLSX/XLS/CSV executor used by analysis.
    ///
    /// # Errors
    ///
    /// Returns when no path is bound or the extension is unsupported.
    pub fn choice_excel_executor(&mut self) -> Result<()> {
        let path = self.path.as_ref().ok_or_else(|| {
            ExcelError::Format(
                "ExcelAnalyserImpl.choiceExcelExecutor requires a workbook path".to_owned(),
            )
        })?;
        let excel_type = detect_excel_type(path)?;
        self.excel_read_executor = Some(ExcelReadExecutorKind::new(
            excel_type,
            path,
            self.options.clone(),
        )?);
        self.excel_type = Some(excel_type);
        // Seed context the way DefaultXlsx/Xls/CsvReadContext would after construction.
        let sheet_hint = match &self.options.sheet {
            crate::SheetSelector::Name(name) => name.clone(),
            _ => String::new(),
        };
        self.context = AnalysisContext::new(sheet_hint, 0, 0)
            .with_custom_object(self.options.custom_object.clone());
        self.last_error = None;
        Ok(())
    }

    /// Java `analysis(List<ReadSheet>, Boolean)` typed hot path.
    ///
    /// Delegates to [`read_xlsx`] / [`read_xls`] / [`read_csv`] based on the
    /// executor chosen by [`choice_excel_executor`]. Sheet selection follows
    /// [`ReadOptions::sheet`] (Java `readAll` / `parameterSheetDataList`).
    ///
    /// # Errors
    ///
    /// Propagates workbook, sheet-selection, conversion, or listener errors.
    /// Also fails when the analyser was already finished or has no path.
    pub fn analysis_with_listener<T, L>(&mut self, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        if self.finished {
            return Err(ExcelError::Unsupported(
                "ExcelAnalyserImpl.analysis called after finish()".to_owned(),
            ));
        }
        if self.excel_type.is_none() {
            self.choice_excel_executor()?;
        }
        if self.path.is_none() {
            return Err(ExcelError::Format(
                "ExcelAnalyserImpl.analysis requires a workbook path".to_owned(),
            ));
        }
        // Java exposes the same mutable AnalysisContext owned by the selected
        // executor. Rust callback contexts are immutable snapshots, so capture
        // every callback before forwarding it and retain the latest snapshot.
        let result = {
            let executor = self.excel_read_executor.as_mut().ok_or_else(|| {
                ExcelError::Format(
                    "ExcelAnalyserImpl.analysis missing ExcelReadExecutor after choiceExcelExecutor"
                        .to_owned(),
                )
            })?;
            let mut tracking_listener = ContextTrackingReadListener {
                delegate: listener,
                latest: &mut self.context,
            };
            executor.execute_with_listener::<T, _>(&self.options, &mut tracking_listener)
        };

        match result {
            Ok(()) => {
                self.last_error = None;
                Ok(())
            }
            Err(error) => {
                // Java analysis() calls finish() on failure before rethrowing.
                self.finish();
                self.last_error = Some(error.clone());
                Err(error)
            }
        }
    }
}

impl ExcelAnalyser for ExcelAnalyserImpl {
    fn analysis<T, L>(&mut self, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        self.analysis_with_listener::<T, L>(listener)
    }

    /// Java `finish()` — releases per-thread caches and temporary input.
    ///
    /// The Rust parsers own workbook/ZIP/CSV handles inside each `execute`
    /// call, so those handles close by RAII before this method runs. Unlike
    /// POI's `Biff8EncryptionKey`, passwords are stored per reader and never
    /// installed in global or thread-local state, so no `clearEncrypt03`
    /// operation is required.
    fn finish(&mut self) {
        if self.finished {
            return;
        }
        crate::read_cache::remove_thread_local_cache();
        easyexcel_core::util::number_data_formatter_utils::remove_thread_local_cache();
        self.temporary_input = None;
        self.finished = true;
    }

    fn excel_executor(&self) -> &ExcelReadExecutorKind {
        self.excel_read_executor
            .as_ref()
            .expect("choice_excel_executor must initialize an executor")
    }

    /// Java `analysisContext()` — always returns a safe context (never `unreachable!`).
    fn analysis_context(&self) -> &AnalysisContext {
        &self.context
    }
}

/// Resolves `ExcelTypeEnum` the way Java `ExcelTypeEnum.valueOf(ReadWorkbook)` does for files.
fn detect_excel_type(path: &Path) -> Result<ExcelTypeEnum> {
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .map(str::to_ascii_lowercase);
    match extension.as_deref() {
        Some("xlsx") | Some("xlsm") => Ok(ExcelTypeEnum::Xlsx),
        Some("xls") => Ok(ExcelTypeEnum::Xls),
        Some("csv") => Ok(ExcelTypeEnum::Csv),
        Some(other) => Err(ExcelError::Format(format!(
            "unsupported excel extension: .{other}"
        ))),
        None => {
            // Java may sniff magic bytes; path-less streams default toward CSV.
            Ok(ExcelTypeEnum::Csv)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use easyexcel_core::DynamicRow;
    use easyexcel_core::DynamicValue;
    use std::io::Write;
    use tempfile::NamedTempFile;

    #[derive(Default)]
    struct CollectingListener {
        rows: Vec<DynamicRow>,
        contexts: Vec<AnalysisContext>,
    }

    impl ReadListener<DynamicRow> for CollectingListener {
        fn invoke(&mut self, data: DynamicRow, context: &AnalysisContext) -> Result<()> {
            self.rows.push(data);
            self.contexts.push(context.clone());
            Ok(())
        }

        fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
            Ok(())
        }
    }

    /// Builds a minimal CSV workbook for analyser dispatch tests.
    fn write_csv() -> NamedTempFile {
        let mut file = NamedTempFile::with_suffix(".csv").expect("temp csv");
        writeln!(file, "name,age").expect("header");
        writeln!(file, "alice,20").expect("row");
        file
    }

    #[test]
    fn analysis_context_is_safe_default_without_path() {
        let analyser = ExcelAnalyserImpl::new();
        let context = analyser.analysis_context();
        assert_eq!(context.sheet_name(), "");
        assert_eq!(context.sheet_no(), 0);
        assert_eq!(context.row_index(), 0);
        assert!(analyser.last_error().is_none());
    }

    #[test]
    fn choice_excel_executor_resolves_csv_extension() -> Result<()> {
        let file = write_csv();
        let analyser = ExcelAnalyserImpl::from_path(file.path(), ReadOptions::default())?;
        assert_eq!(analyser.excel_type(), Some(ExcelTypeEnum::Csv));
        assert!(
            analyser.analysis_context().sheet_name().is_empty()
                || analyser.analysis_context().row_index() == 0
        );
        Ok(())
    }

    #[test]
    fn analysis_with_listener_delegates_to_read_csv() -> Result<()> {
        let file = write_csv();
        let mut options = ReadOptions::default();
        options.head_row_number = 1;
        let mut analyser = ExcelAnalyserImpl::from_path(file.path(), options)?;
        let mut listener = CollectingListener::default();
        analyser.analysis_with_listener::<DynamicRow, _>(&mut listener)?;
        assert_eq!(listener.rows.len(), 1);
        match listener.rows[0].get(0) {
            Some(DynamicValue::String(name)) => assert_eq!(name, "alice"),
            other => panic!("expected alice string cell, got {other:?}"),
        }
        assert_eq!(analyser.analysis_context().sheet_name(), "Sheet1");
        assert_eq!(analyser.analysis_context().sheet_no(), 0);
        assert_eq!(analyser.analysis_context().row_index(), 1);
        assert!(!analyser.is_finished());
        Ok(())
    }

    #[test]
    fn trait_analysis_dispatches_to_the_real_typed_executor() -> Result<()> {
        let file = write_csv();
        let mut options = ReadOptions::default();
        options.head_row_number = 1;
        let mut analyser = ExcelAnalyserImpl::from_path(file.path(), options)?;
        let mut listener = CollectingListener::default();
        ExcelAnalyser::analysis::<DynamicRow, _>(&mut analyser, &mut listener)?;
        assert_eq!(listener.rows.len(), 1);
        assert!(analyser.last_error().is_none());
        Ok(())
    }

    #[test]
    fn finish_is_idempotent() {
        let mut analyser = ExcelAnalyserImpl::new();
        analyser.finish();
        analyser.finish();
        assert!(analyser.is_finished());
    }
}
