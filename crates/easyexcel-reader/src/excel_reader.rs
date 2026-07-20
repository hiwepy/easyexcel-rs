//! Mirrors Java `com.alibaba.excel.ExcelReader`.

use std::marker::PhantomData;
use std::path::PathBuf;

use easyexcel_core::{AnalysisContext, ExcelRow, ReadListener, Result};

use crate::analysis::excel_analyser::ExcelAnalyser;
use crate::analysis::excel_analyser_impl::ExcelAnalyserImpl;
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

    /// Parses every configured worksheet. (Java `readAll()`)
    pub fn read_all(&mut self) -> Result<()> {
        self.analyser.analysis_with_listener(&mut self.listener)
    }

    /// Parses the supplied worksheets. (Java `read(ReadSheet...)`)
    pub fn read(&mut self, sheets: &[ReadSheet]) -> Result<&mut Self> {
        if let Some(last) = sheets.last() {
            self.analyser.set_sheet_selector(if last.sheet_name().is_empty() {
                SheetSelector::Index(last.sheet_no())
            } else {
                SheetSelector::Name(last.sheet_name().to_owned())
            });
        }
        self.analyser.analysis_with_listener(&mut self.listener)?;
        Ok(self)
    }

    /// Returns the live analysis context. (Java `analysisContext()`)
    #[must_use]
    pub fn analysis_context(&self) -> &AnalysisContext {
        ExcelAnalyser::analysis_context(&self.analyser)
    }

    /// Completes the read and releases resources. (Java `finish()`)
    pub fn finish(&mut self) {
        ExcelAnalyser::finish(&mut self.analyser);
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

    use easyexcel_core::DynamicRow;
    use tempfile::NamedTempFile;

    use super::*;
    use crate::ReadOptions;

    #[derive(Default)]
    struct CollectListener {
        rows: Vec<DynamicRow>,
    }

    impl ReadListener<DynamicRow> for CollectListener {
        fn invoke(
            &mut self,
            data: DynamicRow,
            _context: &AnalysisContext,
        ) -> Result<()> {
            self.rows.push(data);
            Ok(())
        }
    }

    #[test]
    fn excel_reader_read_all_loads_csv_rows() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "name,age")?;
        writeln!(file, "reader,30")?;
        let mut listener = CollectListener::default();
        let mut reader = ExcelReader::new(file.path(), ReadOptions::default(), listener)?;
        reader.read_all()?;
        Ok(())
    }
}
