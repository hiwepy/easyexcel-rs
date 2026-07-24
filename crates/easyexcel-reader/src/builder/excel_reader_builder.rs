//! Mirrors Java `com.alibaba.excel.read.builder.ExcelReaderBuilder`.

use std::cell::RefCell;
use std::io::{Read, Write};
use std::path::PathBuf;
use std::rc::Rc;
use std::sync::Arc;

use easyexcel_core::{
    AnalysisContext, CellExtraType, CsvCharset, CustomReadObject, ExcelRow, ReadDefaultReturn,
    ReadListener, ReadListenerList, Result,
};

use crate::builder::abstract_excel_reader_parameter_builder::AbstractExcelReaderParameterBuilder;
use crate::cache::SimpleReadCacheSelector;
use crate::excel_reader::ExcelReader;
use crate::{ReadCacheMode, ReadOptions, SheetSelector, StoredReadCacheSelector};

/// Mirrors Java `ExcelReaderBuilder extends AbstractExcelReaderParameterBuilder`.
#[derive(Debug, Clone, Default)]
pub struct ExcelReaderBuilder {
    /// Mirrors `ReadWorkbook.file`.
    pub file: Option<PathBuf>,
    /// Guard for a caller-supplied Java-style `InputStream`.
    temporary_input: Option<Arc<tempfile::TempPath>>,
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
        self.temporary_input = None;
        self
    }

    /// Materialises a caller-supplied input stream into an automatically
    /// deleted file and selects the existing XLSX/XLS/CSV parsing engine.
    ///
    /// Java accepts a non-seekable `InputStream`, while XLSX and XLS readers
    /// need random access. The temporary file is retained by the resulting
    /// [`ExcelReader`] and deleted only after that reader is dropped.
    pub fn input_stream<R>(mut self, mut input: R) -> Result<Self>
    where
        R: Read,
    {
        let mut bytes = Vec::new();
        input.read_to_end(&mut bytes)?;
        let suffix = input_suffix(&bytes);
        let mut file = tempfile::Builder::new()
            .prefix("easyexcel-input-")
            .suffix(suffix)
            .tempfile()?;
        file.write_all(&bytes)?;
        file.flush()?;
        let temporary_input = Arc::new(file.into_temp_path());
        self.file = Some(temporary_input.to_path_buf());
        self.temporary_input = Some(temporary_input);
        Ok(self)
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

    /// Sets the character encoding used for CSV input. (Java `charset(Charset)`)
    #[must_use]
    pub fn charset(mut self, charset: impl Into<CsvCharset>) -> Self {
        self.options.charset = charset.into();
        self
    }

    /// Controls whether physically empty rows are skipped.
    /// (Java `ignoreEmptyRow(Boolean)`)
    #[must_use]
    pub const fn ignore_empty_row(mut self, ignore: bool) -> Self {
        self.options.ignore_empty_row = ignore;
        self
    }

    /// Stores a value exposed through [`AnalysisContext::custom_object`].
    /// (Java `customObject(Object)`)
    #[must_use]
    pub fn custom_object<C>(mut self, custom_object: C) -> Self
    where
        C: std::any::Any + Send + Sync,
    {
        self.options.custom_object = Some(CustomReadObject::new(custom_object));
        self
    }

    /// Sets the workbook password used by encrypted OOXML input.
    /// (Java `password(String)`)
    #[must_use]
    pub fn password(mut self, password: impl Into<String>) -> Self {
        self.options.password = Some(password.into());
        self
    }

    /// Enables one additional metadata category.
    /// (Java `extraRead(CellExtraTypeEnum)`)
    #[must_use]
    pub fn extra_read(mut self, extra_type: CellExtraType) -> Self {
        self.options.extra_read.insert(extra_type);
        self
    }

    /// Selects the no-model value representation.
    /// (Java `readDefaultReturn(ReadDefaultReturnEnum)`)
    #[must_use]
    pub const fn read_default_return(mut self, mode: ReadDefaultReturn) -> Self {
        self.options.read_default_return = mode;
        self
    }

    /// Controls scientific formatting.
    #[must_use]
    pub fn use_scientific_format(mut self, enabled: bool) -> Self {
        self.options.scientific_format = if enabled {
            crate::ScientificFormatMode::Scientific
        } else {
            crate::ScientificFormatMode::Plain
        };
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
    pub fn simple_read_cache_selector(self, selector: SimpleReadCacheSelector) -> Self {
        self.read_cache_selector(StoredReadCacheSelector::Simple(selector))
    }

    /// Registers the first typed listener and returns a builder that owns it.
    ///
    /// This is the Rust equivalent of Java
    /// `ExcelReaderBuilder.registerReadListener(...)`: the listener becomes
    /// part of the builder state and is passed into the built [`ExcelReader`].
    /// The older [`Self::build`] overload that accepts a listener explicitly
    /// remains available for source compatibility.
    #[must_use]
    pub fn register_read_listener<T, L>(self, listener: L) -> RegisteredExcelReaderBuilder<T>
    where
        T: ExcelRow + Clone,
        L: ReadListener<T> + 'static,
    {
        RegisteredExcelReaderBuilder {
            builder: self,
            listeners: ReadListenerList::new(listener),
        }
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
        match self.temporary_input {
            Some(temporary_input) => {
                ExcelReader::from_temporary_input(path, temporary_input, self.options, listener)
            }
            None => ExcelReader::new(path, self.options, listener),
        }
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

    /// Reads synchronously and returns all converted rows.
    /// (Java `doReadAllSync()`)
    pub fn do_read_all_sync<T>(self) -> Result<Vec<T>>
    where
        T: ExcelRow,
    {
        let rows = Rc::new(RefCell::new(Vec::new()));
        let listener = SharedCollectListener(Rc::clone(&rows));
        let mut reader = self.build(listener)?;
        reader.read_all()?;
        reader.finish();
        drop(reader);
        let collected = std::mem::take(&mut *rows.borrow_mut());
        Ok(collected)
    }
}

fn input_suffix(bytes: &[u8]) -> &'static str {
    const OLE_MAGIC: &[u8; 8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
    if bytes.starts_with(b"PK\x03\x04") {
        ".xlsx"
    } else if bytes.starts_with(OLE_MAGIC) {
        let cursor = std::io::Cursor::new(bytes);
        if cfb::CompoundFile::open(cursor).is_ok_and(|compound| {
            compound.is_stream("/EncryptedPackage") && compound.is_stream("/EncryptionInfo")
        }) {
            ".xlsx"
        } else {
            ".xls"
        }
    } else {
        ".csv"
    }
}

/// An [`ExcelReaderBuilder`] carrying its registered listener.
///
/// Java stores listeners inside `ReadWorkbook`; this wrapper provides the
/// same lifecycle without erasing the Rust row or listener types.
pub struct RegisteredExcelReaderBuilder<T> {
    builder: ExcelReaderBuilder,
    listeners: ReadListenerList<T>,
}

impl<T> RegisteredExcelReaderBuilder<T>
where
    T: ExcelRow + Clone,
{
    /// Sets the file path.
    #[must_use]
    pub fn file(mut self, path: impl Into<PathBuf>) -> Self {
        self.builder = self.builder.file(path);
        self
    }

    /// Selects a worksheet by zero-based index.
    #[must_use]
    pub fn sheet(mut self, index: usize) -> Self {
        self.builder = self.builder.sheet(index);
        self
    }

    /// Selects a worksheet by name.
    #[must_use]
    pub fn sheet_name(mut self, name: impl Into<String>) -> Self {
        self.builder = self.builder.sheet_name(name);
        self
    }

    /// Sets the number of header rows.
    #[must_use]
    pub fn head_row_number(mut self, rows: u32) -> Self {
        self.builder = self.builder.head_row_number(rows);
        self
    }

    /// Sets the CSV character encoding.
    #[must_use]
    pub fn charset(mut self, charset: impl Into<CsvCharset>) -> Self {
        self.builder = self.builder.charset(charset);
        self
    }

    /// Controls whether empty rows are skipped.
    #[must_use]
    pub fn ignore_empty_row(mut self, ignore: bool) -> Self {
        self.builder = self.builder.ignore_empty_row(ignore);
        self
    }

    /// Stores a custom context value.
    #[must_use]
    pub fn custom_object<C>(mut self, custom_object: C) -> Self
    where
        C: std::any::Any + Send + Sync,
    {
        self.builder = self.builder.custom_object(custom_object);
        self
    }

    /// Sets the encrypted OOXML password.
    #[must_use]
    pub fn password(mut self, password: impl Into<String>) -> Self {
        self.builder = self.builder.password(password);
        self
    }

    /// Enables an extra metadata category.
    #[must_use]
    pub fn extra_read(mut self, extra_type: CellExtraType) -> Self {
        self.builder = self.builder.extra_read(extra_type);
        self
    }

    /// Selects the no-model value representation.
    #[must_use]
    pub fn read_default_return(mut self, mode: ReadDefaultReturn) -> Self {
        self.builder = self.builder.read_default_return(mode);
        self
    }

    /// Controls scientific formatting.
    #[must_use]
    pub fn use_scientific_format(mut self, enabled: bool) -> Self {
        self.builder = self.builder.use_scientific_format(enabled);
        self
    }

    /// Registers another listener after all listeners already present.
    #[must_use]
    pub fn register_read_listener<Next>(mut self, listener: Next) -> Self
    where
        Next: ReadListener<T> + 'static,
    {
        self.listeners.push(listener);
        self
    }

    /// Builds an event-driven reader using the registered listener chain.
    pub fn build(self) -> Result<ExcelReader<T, ReadListenerList<T>>> {
        self.builder.build(self.listeners)
    }

    /// Builds, reads, and finishes all configured sheets.
    pub fn do_read_all(self) -> Result<()> {
        self.builder.do_read_all(self.listeners)
    }

    /// Reads synchronously while retaining all previously registered listeners.
    pub fn do_read_all_sync(self) -> Result<Vec<T>> {
        let rows = Rc::new(RefCell::new(Vec::new()));
        let mut collector = SharedCollectListener(Rc::clone(&rows));
        let mut reader = self.build()?;
        reader.read_all_with_additional_listener(&mut collector)?;
        reader.finish();
        drop(reader);
        let collected = std::mem::take(&mut *rows.borrow_mut());
        Ok(collected)
    }
}

struct SharedCollectListener<T>(Rc<RefCell<Vec<T>>>);

impl<T> ReadListener<T> for SharedCollectListener<T> {
    fn invoke(&mut self, data: T, _context: &AnalysisContext) -> Result<()> {
        self.0.borrow_mut().push(data);
        Ok(())
    }
}

impl<T> AbstractExcelReaderParameterBuilder<T> for RegisteredExcelReaderBuilder<T>
where
    T: ExcelRow + Clone,
{
    fn head_row_number(&mut self, head_row_number: i32) -> &mut Self {
        self.builder.options.head_row_number = head_row_number.max(0) as u32;
        self
    }

    fn use_scientific_format(&mut self, enabled: bool) -> &mut Self {
        self.builder.options.scientific_format = if enabled {
            crate::ScientificFormatMode::Scientific
        } else {
            crate::ScientificFormatMode::Plain
        };
        self
    }

    fn register_read_listener(&mut self, listener: Box<dyn ReadListener<T>>) -> &mut Self {
        self.listeners.push_boxed(listener);
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cache::EternalReadCacheSelector;
    use easyexcel_core::DynamicRow;
    use std::io::Write;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};
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

    struct CountingListener(Arc<AtomicUsize>);

    impl ReadListener<DynamicRow> for CountingListener {
        fn invoke(
            &mut self,
            _data: DynamicRow,
            _context: &easyexcel_core::AnalysisContext,
        ) -> Result<()> {
            self.0.fetch_add(1, Ordering::SeqCst);
            Ok(())
        }
    }

    #[test]
    fn builder_applies_eternal_cache_selector_via_read_cache() {
        let builder = ExcelReaderBuilder::new().read_cache_selector(
            StoredReadCacheSelector::Eternal(EternalReadCacheSelector::map_cache()),
        );
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

    #[test]
    fn registered_listeners_are_owned_and_invoked_by_the_builder() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "name,age")?;
        writeln!(file, "bob,22")?;
        let first = Arc::new(AtomicUsize::new(0));
        let second = Arc::new(AtomicUsize::new(0));

        let mut builder = ExcelReaderBuilder::new()
            .file(file.path())
            .register_read_listener::<DynamicRow, _>(CountingListener(Arc::clone(&first)));
        AbstractExcelReaderParameterBuilder::<DynamicRow>::register_read_listener(
            &mut builder,
            Box::new(CountingListener(Arc::clone(&second))),
        );
        builder.do_read_all()?;

        assert_eq!(first.load(Ordering::SeqCst), 1);
        assert_eq!(second.load(Ordering::SeqCst), 1);
        Ok(())
    }

    #[test]
    fn builder_do_read_all_sync_returns_rows() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "name,age")?;
        writeln!(file, "bob,22")?;

        let rows = ExcelReaderBuilder::new()
            .file(file.path())
            .do_read_all_sync::<DynamicRow>()?;

        assert_eq!(rows.len(), 1);
        assert_eq!(
            rows[0].get(0),
            Some(&easyexcel_core::DynamicValue::String("bob".to_owned()))
        );
        Ok(())
    }

    #[test]
    fn registered_builder_sync_read_keeps_existing_listeners() -> Result<()> {
        let mut file = NamedTempFile::with_suffix(".csv")?;
        writeln!(file, "name,age")?;
        writeln!(file, "bob,22")?;
        let invoked = Arc::new(AtomicUsize::new(0));

        let rows = ExcelReaderBuilder::new()
            .file(file.path())
            .register_read_listener::<DynamicRow, _>(CountingListener(Arc::clone(&invoked)))
            .do_read_all_sync()?;

        assert_eq!(invoked.load(Ordering::SeqCst), 1);
        assert_eq!(rows.len(), 1);
        Ok(())
    }

    #[test]
    fn builder_stores_java_workbook_options_in_real_read_options() {
        let builder = ExcelReaderBuilder::new()
            .charset("gbk")
            .ignore_empty_row(false)
            .custom_object(42_u32)
            .password("secret")
            .extra_read(CellExtraType::Comment)
            .read_default_return(ReadDefaultReturn::ActualData);

        assert!(!builder.options.ignore_empty_row);
        assert_eq!(builder.options.password.as_deref(), Some("secret"));
        assert!(builder.options.custom_object.is_some());
        assert!(builder.options.extra_read.contains(&CellExtraType::Comment));
        assert_eq!(
            builder.options.read_default_return,
            ReadDefaultReturn::ActualData
        );
    }
}
