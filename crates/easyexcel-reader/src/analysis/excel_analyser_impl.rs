//! Mirrors Java `com.alibaba.excel.analysis.ExcelAnalyserImpl`.

use std::path::{Path, PathBuf};

use easyexcel_core::support::ExcelTypeEnum;
use easyexcel_core::{AnalysisContext, ExcelError, ExcelRow, ReadListener, Result};

use crate::analysis::excel_analyser::ExcelAnalyser;
use crate::{ReadOptions, SheetSelector, read_csv, read_xls, read_xlsx};

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
    /// Live analysis context. (Java `ExcelAnalyserImpl.analysisContext`)
    context: AnalysisContext,
    /// Prevents multiple shutdowns. (Java `ExcelAnalyserImpl.finished`)
    finished: bool,
    /// Captures errors from the void [`ExcelAnalyser::analysis`] entry.
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
            // Safe default — Java leaves `analysisContext` null until
            // `choiceExcelExecutor`; Rust always exposes a usable value.
            context: AnalysisContext::new("", 0, 0),
            finished: false,
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
            context: AnalysisContext::new("", 0, 0),
            finished: false,
            last_error: None,
        };
        analyser.choice_excel_executor()?;
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
    /// and seeds [`analysis_context`](ExcelAnalyser::analysis_context). The
    /// concrete SAX/CSV executors are not constructed here — Rust dispatches
    /// to [`read_xlsx`] / [`read_xls`] / [`read_csv`] at analysis time.
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
        let path = self.path.as_ref().ok_or_else(|| {
            ExcelError::Format(
                "ExcelAnalyserImpl.analysis requires a workbook path".to_owned(),
            )
        })?;
        let excel_type = self.excel_type.ok_or_else(|| {
            ExcelError::Format(
                "ExcelAnalyserImpl.analysis missing ExcelTypeEnum after choiceExcelExecutor"
                    .to_owned(),
            )
        })?;

        // Java: excelReadExecutor.execute() inside analysis(...).
        let result = match excel_type {
            ExcelTypeEnum::Xlsx => read_xlsx::<T, L>(path, &self.options, listener),
            ExcelTypeEnum::Xls => read_xls::<T, L>(path, &self.options, listener),
            ExcelTypeEnum::Csv => read_csv::<T, L>(path, &self.options, listener),
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
    /// Java `analysis(List<ReadSheet>, Boolean)`.
    ///
    /// The Java signature carries listeners on `ReadWorkbook`; Rust listeners
    /// are passed explicitly via
    /// [`analysis_with_listener`](ExcelAnalyserImpl::analysis_with_listener).
    /// This void entry records a clear [`ExcelError`] instead of panicking.
    fn analysis(&mut self) {
        if self.finished {
            self.last_error = Some(ExcelError::Unsupported(
                "ExcelAnalyserImpl.analysis called after finish()".to_owned(),
            ));
            return;
        }
        if self.path.is_none() {
            self.last_error = Some(ExcelError::Format(
                "ExcelAnalyserImpl.analysis requires a workbook path".to_owned(),
            ));
            return;
        }
        // Typed rows + listener are required for read_* dispatch.
        self.last_error = Some(ExcelError::Unsupported(
            "use ExcelAnalyserImpl::analysis_with_listener::<T, L>(...) to run read_xlsx/read_xls/read_csv"
                .to_owned(),
        ));
    }

    /// Java `finish()` — marks the analyser closed (idempotent).
    fn finish(&mut self) {
        if self.finished {
            return;
        }
        self.finished = true;
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
        assert!(analyser.analysis_context().sheet_name().is_empty()
            || analyser.analysis_context().row_index() == 0);
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
        assert!(!analyser.is_finished());
        Ok(())
    }

    #[test]
    fn void_analysis_records_error_instead_of_panicking() {
        let mut analyser = ExcelAnalyserImpl::new();
        analyser.analysis();
        assert!(analyser.last_error().is_some());
        // Still returns a usable context.
        let _ = analyser.analysis_context();
    }

    #[test]
    fn finish_is_idempotent() {
        let mut analyser = ExcelAnalyserImpl::new();
        analyser.finish();
        analyser.finish();
        assert!(analyser.is_finished());
    }
}
