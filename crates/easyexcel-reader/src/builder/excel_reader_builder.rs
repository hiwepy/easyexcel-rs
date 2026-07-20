//! Mirrors Java `com.alibaba.excel.read.builder.ExcelReaderBuilder`.

use std::path::PathBuf;

use easyexcel_core::{ExcelRow, ReadListener, Result};

use crate::builder::abstract_excel_reader_parameter_builder::AbstractExcelReaderParameterBuilder;
use crate::cache::{EternalReadCacheSelector, SimpleReadCacheSelector};
use crate::excel_reader::ExcelReader;
use crate::{ReadCacheMode, ReadOptions, SheetSelector, StoredReadCacheSelector};

/// Mirrors Java `ExcelReaderBuilder extends AbstractExcelReaderParameterBuilder`.
#[derive(Debug, Clone, Default)]
pub struct ExcelReaderBuilder {
    /// Mirrors `ReadWorkbook.file`.
    pub file: Option<PathBuf>,
    /// Collapsed read options from Java `ReadWorkbook` + parameter builders.
    pub options: ReadOptions,
}

impl ExcelReaderBuilder {
    /// Creates a builder. (Java `new ExcelReaderBuilder()`)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Sets the file path. (Java `file(String pathName)`)
    #[must_use]
    pub fn file(mut self, path: impl Into<PathBuf>) -> Self {
        self.file = Some(path.into());
        self
    }

    /// Selects a worksheet by zero-based index. (Java `sheet(Integer)`)
    #[must_use]
    pub fn sheet(mut self, index: usize) -> Self {
        self.options.sheet = SheetSelector::Index(index);
        self
    }

    /// Selects a worksheet by name. (Java `sheet(String)`)
    #[must_use]
    pub fn sheet_name(mut self, name: impl Into<String>) -> Self {
        self.options.sheet = SheetSelector::Name(name.into());
        self
    }

    /// Sets the number of header rows. (Java `headRowNumber(Integer)`)
    #[must_use]
    pub const fn head_row_number(mut self, rows: u32) -> Self {
        self.options.head_row_number = rows;
        self
    }

    /// Sets the shared-string cache mode directly. (Java `readCache(ReadCache)`)
    #[must_use]
    pub fn read_cache(mut self, mode: ReadCacheMode) -> Self {
        self.options.read_cache = mode;
        self.options.read_cache_selector = None;
        self
    }

    /// Installs a cache selector. (Java `readCacheSelector(ReadCacheSelector)`)
    #[must_use]
    pub fn read_cache_selector(mut self, selector: StoredReadCacheSelector) -> Self {
        self.options.read_cache_selector = Some(selector);
        self
    }

    /// Installs Java's default simple selector.
    #[must_use]
    pub fn simple_read_cache_selector(mut self, selector: SimpleReadCacheSelector) -> Self {
        self.read_cache_selector(StoredReadCacheSelector::Simple(selector))
    }

    /// Builds an event-driven reader. (Java `build()`)
    pub fn build<T, L>(self, listener: L) -> Result<ExcelReader<T, L>>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        let path = self.file.ok_or_else(|| {
            easyexcel_core::ExcelError::Format(
                "ExcelReaderBuilder.file must be set before build()".to_owned(),
            )
        })?;
        ExcelReader::new(path, self.options, listener)
    }

    /// Builds and immediately reads all configured sheets. (Java `doReadAll()`)
    pub fn do_read_all<T, L>(self, listener: L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        let mut reader = self.build(listener)?;
        reader.read_all()
    }
}

impl AbstractExcelReaderParameterBuilder for ExcelReaderBuilder {
    fn head_row_number(&mut self, head_row_number: i32) -> &mut Self {
        self.options.head_row_number = head_row_number.max(0) as u32;
        self
    }

    fn use_scientific_format(&mut self, enabled: bool) -> &mut Self {
        self.options.scientific_format = if enabled {
            crate::ScientificFormatMode::Scientific
        } else {
            crate::ScientificFormatMode::Plain
        };
        self
    }

    fn register_read_listener<T>(
        &mut self,
        _listener: Box<dyn ReadListener<T>>,
    ) -> &mut Self
    {
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use easyexcel_core::DynamicRow;
    use std::io::Write;
    use tempfile::NamedTempFile;

    #[derive(Default)]
    struct CollectListener {
        rows: Vec<DynamicRow>,
    }

    impl ReadListener<DynamicRow> for CollectListener {
        fn invoke(
            &mut self,
            data: DynamicRow,
            _context: &easyexcel_core::AnalysisContext,
        ) -> Result<()> {
            self.rows.push(data);
            Ok(())
        }
    }

    #[test]
    fn builder_applies_eternal_cache_selector_via_read_cache() {
        let builder = ExcelReaderBuilder::new()
            .read_cache_selector(StoredReadCacheSelector::Eternal(
                EternalReadCacheSelector::map_cache(),
            ));
        assert!(builder.options.read_cache_selector.is_some());
    }

    #[test]
    fn builder_reads_csv_file() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "name,age")?;
        writeln!(file, "bob,22")?;
        ExcelReaderBuilder::new()
            .file(file.path())
            .head_row_number(1)
            .do_read_all(CollectListener::default())?;
        Ok(())
    }
}
