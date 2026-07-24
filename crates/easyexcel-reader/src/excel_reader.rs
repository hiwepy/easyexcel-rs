//! Mirrors Java `com.alibaba.excel.ExcelReader`.

use std::marker::PhantomData;
use std::path::PathBuf;

use easyexcel_core::support::ExcelTypeEnum;
use easyexcel_core::{
    AnalysisContext, CompositeReadListener, ExcelError, ExcelRow, ReadListener, Result,
};

use crate::analysis::excel_analyser::ExcelAnalyser;
use crate::analysis::excel_analyser_impl::ExcelAnalyserImpl;
use crate::analysis::excel_read_executor::ExcelReadExecutorKind;
use crate::context::read_sheet::ReadSheet;
use crate::{ReadOptions, SheetSelector};

/// Event-driven workbook reader.
///
/// Mirrors Java `com.alibaba.excel.ExcelReader`.
pub struct ExcelReader<T, L> {
    analyser: ExcelAnalyserImpl,
    listener: L,
    marker: PhantomData<T>,
}

impl<T, L> ExcelReader<T, L>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    /// Creates a reader bound to a workbook path and options.
    ///
    /// Mirrors Java `ExcelReader(ReadWorkbook)`.
    pub fn new(path: impl Into<PathBuf>, options: ReadOptions, listener: L) -> Result<Self> {
        Ok(Self {
            analyser: ExcelAnalyserImpl::from_path(path, options)?,
            listener,
            marker: PhantomData,
        })
    }

    /// Creates a reader for a path owned by a temporary-input guard.
    ///
    /// The compatible builder uses this for Java `read(InputStream, ...)`.
    pub(crate) fn from_temporary_input(
        path: impl Into<PathBuf>,
        temporary_input: std::sync::Arc<tempfile::TempPath>,
        options: ReadOptions,
        listener: L,
    ) -> Result<Self> {
        Ok(Self {
            analyser: ExcelAnalyserImpl::from_temporary_input(path, temporary_input, options)?,
            listener,
            marker: PhantomData,
        })
    }

    /// Returns whether this reader owns a materialised input-stream guard.
    #[must_use]
    pub const fn has_temporary_input(&self) -> bool {
        self.analyser.has_temporary_input()
    }

    /// Parses every configured worksheet. (Java `readAll()`)
    pub fn read_all(&mut self) -> Result<()> {
        ExcelAnalyser::analysis::<T, L>(&mut self.analyser, &mut self.listener)
    }

    /// Deprecated Java `read()` alias for [`Self::read_all`].
    #[deprecated(note = "please use read_all()")]
    pub fn read_deprecated(&mut self) -> Result<()> {
        self.read_all()
    }

    pub(crate) fn read_all_with_additional_listener<M>(&mut self, listener: &mut M) -> Result<()>
    where
        T: Clone,
        M: ReadListener<T>,
    {
        let mut listeners = CompositeReadListener::new(&mut self.listener, listener);
        ExcelAnalyser::analysis::<T, _>(&mut self.analyser, &mut listeners)
    }

    /// Parses the supplied worksheets. (Java `read(ReadSheet...)`)
    pub fn read(&mut self, sheets: &[ReadSheet]) -> Result<&mut Self> {
        Self::read_sheets_with_listener(&mut self.analyser, &mut self.listener, sheets)?;
        Ok(self)
    }

    pub(crate) fn read_with_additional_listener<M>(
        &mut self,
        sheets: &[ReadSheet],
        listener: &mut M,
    ) -> Result<&mut Self>
    where
        T: Clone,
        M: ReadListener<T>,
    {
        let mut listeners = CompositeReadListener::new(&mut self.listener, listener);
        Self::read_sheets_with_listener(&mut self.analyser, &mut listeners, sheets)?;
        Ok(self)
    }

    fn read_sheets_with_listener<M>(
        analyser: &mut ExcelAnalyserImpl,
        listener: &mut M,
        sheets: &[ReadSheet],
    ) -> Result<()>
    where
        M: ReadListener<T>,
    {
        if sheets.is_empty() {
            return Err(ExcelError::Format(
                "Specify at least one read sheet.".to_owned(),
            ));
        }

        let workbook_head_row_number = analyser.options().head_row_number;
        let workbook_scientific_format = analyser.options().scientific_format;
        let path = analyser
            .path()
            .ok_or_else(|| ExcelError::Format("ExcelReader has no workbook path".to_owned()))?;
        let actual_sheets = match analyser.excel_type() {
            Some(ExcelTypeEnum::Xlsx) => crate::list_xlsx_sheets(path, analyser.options())?,
            Some(ExcelTypeEnum::Xls) => crate::list_xls_sheets(path, analyser.options())?,
            Some(ExcelTypeEnum::Csv) => vec![(0, "Sheet1".to_owned())],
            None => {
                return Err(ExcelError::Format(
                    "ExcelReader has no resolved workbook type".to_owned(),
                ));
            }
        };

        // Java executors enumerate actual workbook sheets and use the first
        // matching parameter sheet. This preserves workbook order, ignores
        // duplicate parameters, and leaves unknown selections unread.
        for (actual_sheet_no, actual_sheet_name) in actual_sheets {
            let Some(sheet) = sheets.iter().find(|sheet| {
                (sheet.has_sheet_no() && sheet.sheet_no() == actual_sheet_no)
                    || (!sheet.sheet_name().is_empty()
                        && crate::sheet_name_matches(
                            &actual_sheet_name,
                            sheet.sheet_name(),
                            analyser.options().auto_trim,
                        ))
            }) else {
                continue;
            };
            analyser.set_sheet_selector(SheetSelector::Index(actual_sheet_no));
            let options = analyser.options_mut();
            options.head_row_number = sheet.head_row_number().unwrap_or(workbook_head_row_number);
            options.scientific_format =
                sheet
                    .use_scientific_format()
                    .map_or(workbook_scientific_format, |enabled| {
                        if enabled {
                            crate::ScientificFormatMode::Scientific
                        } else {
                            crate::ScientificFormatMode::Plain
                        }
                    });
            ExcelAnalyser::analysis::<T, _>(analyser, listener)?;
        }
        Ok(())
    }

    /// Returns the live analysis context. (Java `analysisContext()`)
    #[must_use]
    pub fn analysis_context(&self) -> &AnalysisContext {
        ExcelAnalyser::analysis_context(&self.analyser)
    }

    /// Deprecated Java `getAnalysisContext()` alias.
    #[deprecated(note = "please use analysis_context()")]
    #[must_use]
    pub fn get_analysis_context(&self) -> &AnalysisContext {
        self.analysis_context()
    }

    /// Returns the selected XLSX/XLS/CSV executor.
    #[must_use]
    pub fn excel_executor(&self) -> &ExcelReadExecutorKind {
        ExcelAnalyser::excel_executor(&self.analyser)
    }

    /// Completes the read and releases resources. (Java `finish()`)
    pub fn finish(&mut self) {
        ExcelAnalyser::finish(&mut self.analyser);
    }

    /// Java `Closeable.close()` alias. Finishing is idempotent.
    pub fn close(&mut self) {
        self.finish();
    }
}

impl<T, L> Drop for ExcelReader<T, L> {
    fn drop(&mut self) {
        ExcelAnalyser::finish(&mut self.analyser);
    }
}

#[cfg(test)]
mod tests {
    use std::io::Write;
    use std::sync::{Arc, Mutex};

    use easyexcel_core::DynamicRow;
    use rust_xlsxwriter::Workbook;
    use tempfile::NamedTempFile;

    use super::*;
    use crate::ReadOptions;
    #[derive(Default)]
    struct CollectListener {
        rows: Vec<DynamicRow>,
    }

    impl ReadListener<DynamicRow> for CollectListener {
        fn invoke(&mut self, data: DynamicRow, _context: &AnalysisContext) -> Result<()> {
            self.rows.push(data);
            Ok(())
        }
    }

    #[derive(Clone, Default)]
    struct SheetTraceListener {
        sheets: Arc<Mutex<Vec<String>>>,
    }

    impl ReadListener<DynamicRow> for SheetTraceListener {
        fn invoke(&mut self, _data: DynamicRow, context: &AnalysisContext) -> Result<()> {
            self.sheets
                .lock()
                .expect("sheet trace lock")
                .push(context.sheet_name().to_owned());
            Ok(())
        }
    }

    fn multi_sheet_workbook() -> Result<NamedTempFile> {
        let file = NamedTempFile::with_suffix(".xlsx")?;
        let mut workbook = Workbook::new();
        let first = workbook.add_worksheet();
        first
            .set_name("First")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        first
            .write_string(0, 0, "Value")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        first
            .write_string(1, 0, "one")
            .map_err(|error| ExcelError::Format(error.to_string()))?;

        let second = workbook.add_worksheet();
        second
            .set_name("Second")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        second
            .write_string(0, 0, "ignored heading")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        second
            .write_string(1, 0, "Value")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        second
            .write_string(2, 0, "two")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        workbook
            .save(file.path())
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        Ok(file)
    }

    #[test]
    fn excel_reader_read_all_loads_csv_rows() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "name,age")?;
        writeln!(file, "reader,30")?;
        let listener = CollectListener::default();
        let mut reader = ExcelReader::new(file.path(), ReadOptions::default(), listener)?;
        reader.read_all()?;
        Ok(())
    }

    #[test]
    #[allow(deprecated)]
    fn excel_reader_exposes_the_real_executor_and_java_lifecycle_aliases() -> Result<()> {
        let file = multi_sheet_workbook()?;
        let listener = SheetTraceListener::default();
        let mut reader = ExcelReader::new(file.path(), ReadOptions::default(), listener)?;

        let sheets = reader.excel_executor().sheet_list();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].sheet_name(), "First");
        assert_eq!(sheets[1].sheet_name(), "Second");
        assert!(std::ptr::eq(
            reader.analysis_context(),
            reader.get_analysis_context()
        ));

        reader.read_deprecated()?;
        reader.close();
        reader.close();
        let error = reader
            .read_all()
            .expect_err("a closed reader must reject another analysis");
        assert!(error.to_string().contains("called after finish"));
        Ok(())
    }

    #[test]
    fn excel_reader_csv_executor_reports_its_real_logical_sheet() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "value")?;
        let reader = ExcelReader::new(
            file.path(),
            ReadOptions::default(),
            CollectListener::default(),
        )?;
        let sheets = reader.excel_executor().sheet_list();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].sheet_no(), 0);
        assert_eq!(sheets[0].sheet_name(), "Sheet1");
        Ok(())
    }

    #[test]
    fn excel_reader_read_rejects_an_empty_sheet_list_like_java() -> Result<()> {
        let file = multi_sheet_workbook()?;
        let listener = SheetTraceListener::default();
        let mut reader = ExcelReader::new(file.path(), ReadOptions::default(), listener)?;
        let error = match reader.read(&[]) {
            Ok(_) => panic!("empty sheet list must fail"),
            Err(error) => error,
        };
        assert_eq!(
            error.to_string(),
            "excel format error: Specify at least one read sheet."
        );
        Ok(())
    }

    #[test]
    fn excel_reader_read_processes_each_sheet_and_applies_sheet_parameters() -> Result<()> {
        let file = multi_sheet_workbook()?;
        let listener = SheetTraceListener::default();
        let observed = Arc::clone(&listener.sheets);
        let mut reader = ExcelReader::new(file.path(), ReadOptions::default(), listener)?;
        let first = ReadSheet::new(0);
        let mut second = ReadSheet::named("Second");
        second.set_head_row_number(2);

        // Java enumerates actual workbook sheets, not parameter-list order.
        reader.read(&[second, first])?;

        assert_eq!(
            *observed.lock().expect("sheet trace lock"),
            vec!["First".to_owned(), "Second".to_owned()]
        );
        Ok(())
    }
}
