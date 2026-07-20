//! Public facade for typed, event-driven Excel reading and writing.

mod excel_builder;

use std::io::{Read, Write};
use std::marker::PhantomData;
use std::path::{Path, PathBuf};

pub use easyexcel_core::*;
pub use easyexcel_derive::ExcelRow;
pub use easyexcel_core::metadata::GlobalConfiguration;
pub use easyexcel_reader::{
    apply_global_configuration_to_read_options, Ehcache, EternalReadCacheSelector, ExcelLocale,
    ExcelReader, global_configuration_from_read_options, MapCache, ReadCache, ReadCacheMode,
    ReadCacheSelector, SimpleReadCacheSelector, StoredReadCacheSelector, XlsCache,
};
use easyexcel_reader::{
    ReadOptions, ScientificFormatMode, SheetSelector, read_csv, read_xls, read_xlsx,
};
pub use easyexcel_template::{
    ExcelTemplateWriter, FillConfig, FillDirection, FillWrapper, IntoTemplateValue, TemplateData,
    TemplateSheet, fill_xlsx_template, fill_xlsx_template_list,
};
pub use easyexcel_writer::{
    CellStyle, CsvEncodingWriter, ExcelBuilder, ExcelBuilderImpl, ExcelOutputStream,
    ExcelWriter, HorizontalAlignment, HorizontalCellStyleStrategy,
    LongestMatchColumnWidthStyleStrategy, LoopMergeStrategy, MergeRange,
    SimpleColumnWidthStyleStrategy, SimpleRowHeightStyleStrategy, VerticalAlignment,
    VerticalCellStyleStrategy, WriteOptions, WriteSheet, write_csv_to_buffer,
    write_csv_to_writer, write_xls, write_xls_to_writer, write_xlsx_to_writer,
};
pub use excel_builder::{
    builder_from_writer, do_fill_template, do_fill_template_with_config,
    fill_builder_from_writer, wire_template_fill,
};
use easyexcel_writer::{
    write_csv_with_handlers, write_xls_with_handlers, write_xlsx_with_handlers,
};

/// Static factory matching Java `EasyExcel`'s entry point.
pub struct EasyExcel;

/// Java-compatible alias for [`EasyExcel`].
///
/// Mirrors Java `EasyExcelFactory`; `EasyExcel` extends the same factory in Java.
pub type EasyExcelFactory = EasyExcel;

impl EasyExcel {
    /// Starts an event-driven XLSX, XLS, or CSV read selected from the path extension.
    pub fn read<T, L>(path: impl Into<PathBuf>, listener: L) -> ExcelReaderBuilder<T, L>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        ExcelReaderBuilder {
            path: path.into(),
            options: ReadOptions::default(),
            listener,
            marker: PhantomData,
        }
    }

    /// Starts a synchronous read that collects all converted rows.
    pub fn read_sync<T>(path: impl Into<PathBuf>) -> ExcelSyncReaderBuilder<T>
    where
        T: ExcelRow,
    {
        ExcelSyncReaderBuilder {
            path: path.into(),
            options: ReadOptions::default(),
            marker: PhantomData,
        }
    }

    /// Starts a Java-compatible no-model event read.
    pub fn read_dynamic<L>(
        path: impl Into<PathBuf>,
        listener: L,
    ) -> ExcelReaderBuilder<DynamicRow, L>
    where
        L: ReadListener<DynamicRow>,
    {
        Self::read(path, listener)
    }

    /// Starts a Java-compatible no-model synchronous read.
    #[must_use]
    pub fn read_dynamic_sync(path: impl Into<PathBuf>) -> ExcelSyncReaderBuilder<DynamicRow> {
        Self::read_sync(path)
    }

    /// Starts a new XLSX or CSV write, selected from the path extension.
    pub fn write<T>(path: impl Into<PathBuf>) -> ExcelWriterBuilder<T>
    where
        T: ExcelRow,
    {
        ExcelWriterBuilder {
            path: path.into(),
            options: WriteOptions::default(),
            handlers: Vec::new(),
            marker: PhantomData,
        }
    }

    /// Creates typed worksheet metadata for a stateful [`ExcelWriter`].
    #[must_use]
    pub fn writer_sheet<T>(name: impl Into<String>) -> WriteSheet<T>
    where
        T: ExcelRow,
    {
        WriteSheet::new(name)
    }

    /// Creates typed worksheet metadata for a Java-style zero-based sheet number.
    #[must_use]
    pub fn writer_sheet_index<T>(index: usize) -> WriteSheet<T>
    where
        T: ExcelRow,
    {
        WriteSheet::new_index(index)
    }

    /// Creates a `WriteTable` value mirroring Java
    /// `EasyExcelFactory.writerTable(Integer)`. (Java `writerTable(int)`)
    #[must_use]
    pub fn writer_table(table_no: i32) -> easyexcel_writer::MirroredWriteTable {
        easyexcel_writer::MirroredWriteTable::with_table_no(table_no)
    }

    /// Begins a multi-table write flow that produces an `ExcelWriterTableBuilder`.
    ///
    /// Mirrors Java `ExcelWriterBuilder.table(Integer)` which yields an
    /// `ExcelWriterTableBuilder` for configuring per-table options before
    /// calling `.do_write(rows, sheet, table)`.
    ///
    /// Phase 4 addition: provides the three-arg `write(Collection, WriteSheet, WriteTable)`
    /// overload at the public facade level.
    #[must_use]
    pub fn writer_table_builder(table_no: i32) -> easyexcel_writer::ExcelWriterTableBuilder {
        easyexcel_writer::ExcelWriterTableBuilder::new().table_no(table_no)
    }

    /// Fills scalar `{key}` placeholders in an existing XLSX template.
    ///
    /// Legacy `.xls` templates return typed [`ExcelError::Unsupported`]
    /// (`legacy XLS template fill is not supported`). Java maps this to
    /// `ExcelWriter.fill` on `HSSFWorkbook`; Rust fill remains OOXML-only.
    /// Use [`Self::write`] / `with_template` for `.xls` cell append instead.
    ///
    /// # Errors
    ///
    /// Returns an I/O, Unsupported, or package format error.
    pub fn fill_template(
        template: impl AsRef<Path>,
        output: impl AsRef<Path>,
        data: &TemplateData,
    ) -> Result<()> {
        fill_xlsx_template(template.as_ref(), output.as_ref(), data)
    }

    /// Expands a collection in an existing XLSX template.
    ///
    /// XLS (`.xls`) collection fill is not supported — returns
    /// [`ExcelError::Unsupported`] with `legacy XLS template fill is not supported`.
    ///
    /// # Errors
    ///
    /// Returns an I/O, Unsupported, or OOXML package error.
    pub fn fill_template_list(
        template: impl AsRef<Path>,
        output: impl AsRef<Path>,
        data: &FillWrapper,
        config: FillConfig,
    ) -> Result<()> {
        fill_xlsx_template_list(template.as_ref(), output.as_ref(), data, config)
    }

    /// Loads an XLSX template for repeated Java-style `fill` calls.
    ///
    /// XLS (`.xls`) stateful template writers are rejected with
    /// `legacy XLS template fill is not supported` (use [`Self::fill_template`] for
    /// scalar `.xls` fill).
    ///
    /// # Errors
    ///
    /// Returns an I/O, Unsupported, or OOXML package error when the template cannot be read.
    pub fn template_writer(
        template: impl AsRef<Path>,
        output: impl Into<PathBuf>,
    ) -> Result<ExcelTemplateWriter<'static>> {
        ExcelTemplateWriter::new(template, output)
    }

    /// Loads an XLSX template from a Java-style input stream and writes to a path.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn template_writer_from_reader<R>(
        template: R,
        output: impl Into<PathBuf>,
    ) -> Result<ExcelTemplateWriter<'static>>
    where
        R: Read,
    {
        ExcelTemplateWriter::from_reader(template, output)
    }

    /// Loads a path template and writes to a caller-owned output stream.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn template_writer_to_writer<W>(
        template: impl AsRef<Path>,
        output: &mut W,
    ) -> Result<ExcelTemplateWriter<'_>>
    where
        W: Write,
    {
        ExcelTemplateWriter::to_writer(template, output)
    }

    /// Loads a stream template and writes to a caller-owned output stream.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn template_writer_from_reader_to_writer<R, W>(
        template: R,
        output: &mut W,
    ) -> Result<ExcelTemplateWriter<'_>>
    where
        R: Read,
        W: Write,
    {
        ExcelTemplateWriter::from_reader_to_writer(template, output)
    }

    /// Loads a path template and writes to a closeable output stream.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn template_writer_to_output_stream<'a, W>(
        template: impl AsRef<Path>,
        output: ExcelOutputStream<W>,
    ) -> Result<ExcelTemplateWriter<'a>>
    where
        W: Write + 'a,
    {
        ExcelTemplateWriter::to_output_stream(template, output)
    }

    /// Loads a stream template and writes to a closeable output stream.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn template_writer_from_reader_to_output_stream<'a, R, W>(
        template: R,
        output: ExcelOutputStream<W>,
    ) -> Result<ExcelTemplateWriter<'a>>
    where
        R: Read,
        W: Write + 'a,
    {
        ExcelTemplateWriter::from_reader_to_output_stream(template, output)
    }
}

/// Input accepted by `.sheet(...)`.
pub trait IntoSheetSelector {
    /// Converts to internal sheet selection.
    fn into_sheet_selector(self) -> SheetSelector;
}

impl IntoSheetSelector for usize {
    fn into_sheet_selector(self) -> SheetSelector {
        SheetSelector::Index(self)
    }
}

impl IntoSheetSelector for &str {
    fn into_sheet_selector(self) -> SheetSelector {
        SheetSelector::Name(self.to_owned())
    }
}

impl IntoSheetSelector for String {
    fn into_sheet_selector(self) -> SheetSelector {
        SheetSelector::Name(self)
    }
}

/// Event-driven reader builder.
pub struct ExcelReaderBuilder<T, L> {
    path: PathBuf,
    options: ReadOptions,
    listener: L,
    marker: PhantomData<T>,
}

impl<T, L> ExcelReaderBuilder<T, L>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    /// Selects a worksheet by name or zero-based index.
    #[must_use]
    pub fn sheet(mut self, sheet: impl IntoSheetSelector) -> Self {
        self.options.sheet = sheet.into_sheet_selector();
        self
    }

    /// Selects every worksheet in workbook order.
    #[must_use]
    pub fn all_sheets(mut self) -> Self {
        self.options.sheet = SheetSelector::All;
        self
    }

    /// Sets the number of header rows.
    #[must_use]
    pub const fn head_row_number(mut self, rows: u32) -> Self {
        self.options.head_row_number = rows;
        self
    }

    /// Configures empty-row filtering.
    #[must_use]
    pub const fn ignore_empty_row(mut self, ignore: bool) -> Self {
        self.options.ignore_empty_row = ignore;
        self
    }

    /// Enables or disables Java EasyExcel-compatible string trimming.
    #[must_use]
    pub const fn auto_trim(mut self, enabled: bool) -> Self {
        self.options.auto_trim = enabled;
        self
    }

    /// Selects Excel's 1904 date windowing system for numeric date cells.
    #[must_use]
    pub const fn use_1904_windowing(mut self, enabled: bool) -> Self {
        self.options.use_1904_windowing = enabled;
        self
    }

    /// Controls scientific notation for extreme General-format numeric cells.
    #[must_use]
    pub const fn use_scientific_format(mut self, enabled: bool) -> Self {
        self.options.scientific_format = if enabled {
            ScientificFormatMode::Scientific
        } else {
            ScientificFormatMode::Plain
        };
        self
    }

    /// Sets the locale used for formatted number and date display values.
    #[must_use]
    pub fn locale(mut self, locale: ExcelLocale) -> Self {
        self.options.locale = locale;
        self
    }

    /// Registers a Java-style global converter for this read operation.
    #[must_use]
    pub fn register_converter<V, C>(mut self, converter: C) -> Self
    where
        V: 'static,
        C: Converter<V> + Send + Sync + 'static,
    {
        self.options.converters.register::<V, C>(converter);
        self
    }

    /// Selects the XLSX shared-string cache backend.
    #[must_use]
    pub fn read_cache(mut self, mode: ReadCacheMode) -> Self {
        self.options.read_cache = mode;
        self.options.read_cache_selector = None;
        self
    }

    /// Installs a Java-style cache selector. (Java `readCacheSelector(ReadCacheSelector)`)
    #[must_use]
    pub fn read_cache_selector(mut self, selector: StoredReadCacheSelector) -> Self {
        self.options.read_cache_selector = Some(selector);
        self
    }

    /// Sets the first physical data row to dispatch, zero-based and inclusive.
    ///
    /// Configured header rows are still analysed for name-based mapping.
    #[must_use]
    pub const fn start_row(mut self, row: u32) -> Self {
        self.options.start_row = Some(row);
        self
    }

    /// Sets the last physical data row to dispatch, zero-based and inclusive.
    ///
    /// Configured header rows are still analysed for name-based mapping.
    #[must_use]
    pub const fn end_row(mut self, row: u32) -> Self {
        self.options.end_row = Some(row);
        self
    }

    /// Limits data callbacks to an inclusive physical row range.
    #[must_use]
    pub const fn read_rows(mut self, start: u32, end: u32) -> Self {
        self.options.start_row = Some(start);
        self.options.end_row = Some(end);
        self
    }

    /// Maps a workbook header name to the name used by typed row mapping.
    #[must_use]
    pub fn header_alias(mut self, header: impl Into<String>, alias: impl Into<String>) -> Self {
        self.options
            .header_aliases
            .insert(header.into(), alias.into());
        self
    }

    /// Stores a type-safe value exposed by every read callback context.
    #[must_use]
    pub fn custom_object<C>(mut self, custom_object: C) -> Self
    where
        C: std::any::Any + Send + Sync,
    {
        self.options.custom_object = Some(CustomReadObject::new(custom_object));
        self
    }

    /// Selects the Java-compatible no-model return mode.
    #[must_use]
    pub const fn read_default_return(mut self, mode: ReadDefaultReturn) -> Self {
        self.options.read_default_return = mode;
        self
    }

    /// Enables a Java `extraRead` metadata category.
    #[must_use]
    pub fn extra_read(mut self, extra_type: CellExtraType) -> Self {
        self.options.extra_read.insert(extra_type);
        self
    }

    /// Sets the password for an encrypted OOXML workbook.
    #[must_use]
    pub fn password(mut self, password: impl Into<String>) -> Self {
        self.options.password = Some(password.into());
        self
    }

    /// Sets the character encoding used for CSV input.
    #[must_use]
    pub fn charset(mut self, charset: impl Into<CsvCharset>) -> Self {
        self.options.charset = charset.into();
        self
    }

    /// Executes the read and consumes the builder.
    ///
    /// # Errors
    ///
    /// Returns a workbook, sheet-selection, conversion, or listener error.
    pub fn do_read(mut self) -> Result<()> {
        if is_csv_path(&self.path) {
            read_csv::<T, L>(&self.path, &self.options, &mut self.listener)
        } else if is_xls_path(&self.path) {
            read_xls::<T, L>(&self.path, &self.options, &mut self.listener)
        } else {
            read_xlsx::<T, L>(&self.path, &self.options, &mut self.listener)
        }
    }
}

/// Synchronous collecting reader builder.
pub struct ExcelSyncReaderBuilder<T> {
    path: PathBuf,
    options: ReadOptions,
    marker: PhantomData<T>,
}

impl<T> ExcelSyncReaderBuilder<T>
where
    T: ExcelRow,
{
    /// Selects a worksheet by name or zero-based index.
    #[must_use]
    pub fn sheet(mut self, sheet: impl IntoSheetSelector) -> Self {
        self.options.sheet = sheet.into_sheet_selector();
        self
    }

    /// Selects every worksheet in workbook order.
    #[must_use]
    pub fn all_sheets(mut self) -> Self {
        self.options.sheet = SheetSelector::All;
        self
    }

    /// Sets the number of header rows.
    #[must_use]
    pub const fn head_row_number(mut self, rows: u32) -> Self {
        self.options.head_row_number = rows;
        self
    }

    /// Includes or skips rows containing no values.
    #[must_use]
    pub const fn ignore_empty_row(mut self, ignore: bool) -> Self {
        self.options.ignore_empty_row = ignore;
        self
    }

    /// Enables or disables Java EasyExcel-compatible string trimming.
    #[must_use]
    pub const fn auto_trim(mut self, enabled: bool) -> Self {
        self.options.auto_trim = enabled;
        self
    }

    /// Selects Excel's 1904 date windowing system while collecting rows.
    #[must_use]
    pub const fn use_1904_windowing(mut self, enabled: bool) -> Self {
        self.options.use_1904_windowing = enabled;
        self
    }

    /// Controls scientific notation while collecting extreme General-format numbers.
    #[must_use]
    pub const fn use_scientific_format(mut self, enabled: bool) -> Self {
        self.options.scientific_format = if enabled {
            ScientificFormatMode::Scientific
        } else {
            ScientificFormatMode::Plain
        };
        self
    }

    /// Sets the locale used while collecting formatted number and date values.
    #[must_use]
    pub fn locale(mut self, locale: ExcelLocale) -> Self {
        self.options.locale = locale;
        self
    }

    /// Registers a Java-style global converter while collecting rows.
    #[must_use]
    pub fn register_converter<V, C>(mut self, converter: C) -> Self
    where
        V: 'static,
        C: Converter<V> + Send + Sync + 'static,
    {
        self.options.converters.register::<V, C>(converter);
        self
    }

    /// Selects the XLSX shared-string cache backend while collecting rows.
    #[must_use]
    pub fn read_cache(mut self, mode: ReadCacheMode) -> Self {
        self.options.read_cache = mode;
        self.options.read_cache_selector = None;
        self
    }

    /// Installs a Java-style cache selector while collecting rows.
    #[must_use]
    pub fn read_cache_selector(mut self, selector: StoredReadCacheSelector) -> Self {
        self.options.read_cache_selector = Some(selector);
        self
    }

    /// Sets the first physical data row to collect, zero-based and inclusive.
    ///
    /// Configured header rows are still analysed for name-based mapping.
    #[must_use]
    pub const fn start_row(mut self, row: u32) -> Self {
        self.options.start_row = Some(row);
        self
    }

    /// Sets the last physical data row to collect, zero-based and inclusive.
    ///
    /// Configured header rows are still analysed for name-based mapping.
    #[must_use]
    pub const fn end_row(mut self, row: u32) -> Self {
        self.options.end_row = Some(row);
        self
    }

    /// Limits collected data to an inclusive physical row range.
    #[must_use]
    pub const fn read_rows(mut self, start: u32, end: u32) -> Self {
        self.options.start_row = Some(start);
        self.options.end_row = Some(end);
        self
    }

    /// Maps a workbook header name to the name used by typed row mapping.
    #[must_use]
    pub fn header_alias(mut self, header: impl Into<String>, alias: impl Into<String>) -> Self {
        self.options
            .header_aliases
            .insert(header.into(), alias.into());
        self
    }

    /// Stores a type-safe value exposed while synchronously collecting rows.
    #[must_use]
    pub fn custom_object<C>(mut self, custom_object: C) -> Self
    where
        C: std::any::Any + Send + Sync,
    {
        self.options.custom_object = Some(CustomReadObject::new(custom_object));
        self
    }

    /// Selects the Java-compatible no-model return mode.
    #[must_use]
    pub const fn read_default_return(mut self, mode: ReadDefaultReturn) -> Self {
        self.options.read_default_return = mode;
        self
    }

    /// Enables a Java `extraRead` metadata category.
    #[must_use]
    pub fn extra_read(mut self, extra_type: CellExtraType) -> Self {
        self.options.extra_read.insert(extra_type);
        self
    }

    /// Sets the password for an encrypted OOXML workbook.
    #[must_use]
    pub fn password(mut self, password: impl Into<String>) -> Self {
        self.options.password = Some(password.into());
        self
    }

    /// Sets the character encoding used for CSV input.
    #[must_use]
    pub fn charset(mut self, charset: impl Into<CsvCharset>) -> Self {
        self.options.charset = charset.into();
        self
    }

    /// Reads all rows into memory.
    ///
    /// # Errors
    ///
    /// Returns a workbook, sheet-selection, or row-conversion error.
    pub fn do_read_sync(self) -> Result<Vec<T>> {
        let mut listener = CollectListener(Vec::new());
        if is_csv_path(&self.path) {
            read_csv::<T, _>(&self.path, &self.options, &mut listener)?;
        } else if is_xls_path(&self.path) {
            read_xls::<T, _>(&self.path, &self.options, &mut listener)?;
        } else {
            read_xlsx::<T, _>(&self.path, &self.options, &mut listener)?;
        }
        Ok(listener.0)
    }
}

struct CollectListener<T>(Vec<T>);

impl<T> ReadListener<T> for CollectListener<T> {
    fn invoke(&mut self, data: T, _context: &AnalysisContext) -> Result<()> {
        self.0.push(data);
        Ok(())
    }
}

/// New-workbook writer builder.
pub struct ExcelWriterBuilder<T> {
    path: PathBuf,
    options: WriteOptions,
    handlers: Vec<Box<dyn WriteHandler>>,
    marker: PhantomData<T>,
}

impl<T> ExcelWriterBuilder<T>
where
    T: ExcelRow,
{
    /// Sets the worksheet name.
    #[must_use]
    pub fn sheet(mut self, name: impl Into<String>) -> Self {
        self.options.sheet_name = name.into();
        self
    }

    /// Sets the Java-style zero-based logical worksheet number.
    #[must_use]
    pub fn sheet_index(mut self, index: usize) -> Self {
        self.options.sheet_index = Some(index);
        self.options.sheet_name = index.to_string();
        self
    }

    /// Enables or disables the header row.
    #[must_use]
    pub const fn need_head(mut self, need_head: bool) -> Self {
        self.options.need_head = need_head;
        self
    }

    /// Sets the relative head row index. (Java `ExcelWriterBuilder.relativeHeadRowIndex`)
    ///
    /// When `index > 0`, the header (and subsequent data rows) start at that
    /// zero-based row, leaving the rows above blank — matching Java
    /// `WriteBasicParameter.relativeHeadRowIndex`.
    #[must_use]
    pub const fn relative_head_row_index(mut self, index: i32) -> Self {
        self.options.relative_head_row_index = index;
        self
    }

    /// Freezes the header row.
    #[must_use]
    pub const fn freeze_head(mut self, freeze: bool) -> Self {
        self.options.freeze_head = freeze;
        self
    }

    /// Freezes rows and columns above and to the left of the position.
    #[must_use]
    pub const fn freeze_panes(mut self, row: u32, column: u16) -> Self {
        self.options.freeze_panes = Some((row, column));
        self
    }

    /// Includes only the supplied physical column indexes.
    #[must_use]
    pub fn include_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        self.options.include_column_indexes = Some(indexes.into_iter().collect());
        self
    }

    /// Includes only the supplied Rust field names.
    #[must_use]
    pub fn include_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        self.options.include_column_field_names = Some(names.into_iter().map(Into::into).collect());
        self
    }

    /// Excludes physical column indexes.
    #[must_use]
    pub fn exclude_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        self.options.exclude_column_indexes = indexes.into_iter().collect();
        self
    }

    /// Excludes Rust field names.
    #[must_use]
    pub fn exclude_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        self.options.exclude_column_field_names = names.into_iter().map(Into::into).collect();
        self
    }

    /// Orders selected columns by the corresponding include list.
    #[must_use]
    pub const fn order_by_include_column(mut self, enabled: bool) -> Self {
        self.options.order_by_include_column = enabled;
        self
    }

    /// Adds an absolute merged-cell range using zero-based inclusive coordinates.
    #[must_use]
    pub fn merge_cells(mut self, range: MergeRange) -> Self {
        self.options.merge_ranges.push(range);
        self
    }

    /// Enables automatic width calculation for used columns.
    #[must_use]
    pub const fn auto_width(mut self) -> Self {
        self.options.auto_width = true;
        self
    }

    /// Sets an explicit width for a zero-based physical column.
    #[must_use]
    pub fn column_width(mut self, column: u16, width: u16) -> Self {
        self.options.column_widths.push((column, width));
        self
    }

    /// Replaces the default bold header style.
    #[must_use]
    pub fn head_style(mut self, style: CellStyle) -> Self {
        self.options.head_style = style;
        self
    }

    /// Applies one style to every content row.
    #[must_use]
    pub fn content_style(mut self, style: CellStyle) -> Self {
        self.options.content_styles = vec![style];
        self
    }

    /// Cycles the supplied styles across content rows.
    #[must_use]
    pub fn content_styles(mut self, styles: impl IntoIterator<Item = CellStyle>) -> Self {
        self.options.content_styles = styles.into_iter().collect();
        self
    }

    /// Registers a Java-style global converter for this workbook.
    #[must_use]
    pub fn register_converter<V, C>(mut self, converter: C) -> Self
    where
        V: 'static,
        C: Converter<V> + Send + Sync + 'static,
    {
        self.options.converters.register::<V, C>(converter);
        self
    }

    /// Registers a repeating data-row merge strategy.
    #[must_use]
    pub fn loop_merge(mut self, strategy: LoopMergeStrategy) -> Self {
        self.options.loop_merges.push(strategy);
        self
    }

    /// Replaces derived headers with dynamic multi-level head paths.
    #[must_use]
    pub fn head<S, P>(mut self, paths: impl IntoIterator<Item = P>) -> Self
    where
        S: Into<String>,
        P: IntoIterator<Item = S>,
    {
        self.options.dynamic_head = Some(
            paths
                .into_iter()
                .map(|path| path.into_iter().map(Into::into).collect())
                .collect(),
        );
        self
    }

    /// Registers a write lifecycle handler. Handlers execute by ascending order.
    #[must_use]
    pub fn register_write_handler(mut self, handler: impl WriteHandler + 'static) -> Self {
        self.handlers.push(Box::new(handler));
        self
    }

    /// Encrypts XLSX output using ECMA-376 Agile Encryption.
    #[must_use]
    pub fn password(mut self, password: impl Into<String>) -> Self {
        self.options.password = Some(password.into());
        self
    }

    /// Sets the character encoding used for CSV output.
    #[must_use]
    pub fn charset(mut self, charset: impl Into<CsvCharset>) -> Self {
        self.options.charset = charset.into();
        self
    }

    /// Enables or disables the CSV byte-order mark. Java `EasyExcel` defaults to enabled.
    #[must_use]
    pub const fn with_bom(mut self, enabled: bool) -> Self {
        self.options.with_bom = enabled;
        self
    }

    /// Sets a template workbook file. (Java `ExcelWriterBuilder.withTemplate(String/File)`)
    ///
    /// The template is loaded fully into memory (Java warns this can OOM for large
    /// files). Typed `do_write` / stateful `write` appends after existing template
    /// rows on the selected sheet and keeps other template sheets.
    ///
    /// # Notes
    ///
    /// - CSV templates are rejected (`csv cannot use template.`), matching Java.
    /// - **XLS templates:** record-preserving BIFF8 overlay via
    ///   `easyexcel_writer::biff8::Biff8TemplatePackage` (unmodified records kept;
    ///   new cells appended as LABEL/NUMBER). Creating sheets absent from the
    ///   template remains unsupported.
    /// - **Default (XLSX):** styles and merges are preserved via ZIP/OOXML append
    ///   (`styles.xml` + `mergeCells` kept; new rows appended to `sheetData`).
    ///   Creating a sheet absent from the template adds an empty worksheet part
    ///   without rewriting existing sheets (styles/merges stay intact).
    /// - Images / comments / drawings / column widths from the template remain in
    ///   the package on the ZIP (XLSX) path.
    /// - Opt into value-only replay for XLSX (styles/merges discarded) with
    ///   [`Self::use_legacy_template_seed`].
    #[must_use]
    pub fn with_template(mut self, path: impl Into<PathBuf>) -> Self {
        self.options.template_file = Some(path.into());
        self.options.template_bytes = None;
        self
    }

    /// Sets a template from owned bytes. (Java `ExcelWriterBuilder.withTemplate(InputStream)`)
    ///
    /// Same semantics as [`Self::with_template`]; the stream/file is fully buffered.
    #[must_use]
    pub fn with_template_bytes(mut self, bytes: impl Into<Vec<u8>>) -> Self {
        self.options.template_bytes = Some(bytes.into());
        self.options.template_file = None;
        self
    }

    /// Explicitly enables the legacy calamine → `rust_xlsxwriter` template seed.
    ///
    /// When enabled, `with_template` replays cell **values** only — styles, merges,
    /// images, comments, and drawings are not preserved. Default is `false` (ZIP
    /// preserve). Prefer the default unless you need the legacy seed for debugging.
    #[must_use]
    pub const fn use_legacy_template_seed(mut self, enabled: bool) -> Self {
        self.options.use_legacy_template_seed = enabled;
        self
    }

    /// Redirects this write from its logical path to a caller-owned XLSX stream.
    ///
    /// The path remains available to handler contexts but no file is created.
    /// Borrowing the stream makes ownership explicit and corresponds to Java
    /// `EasyExcel`'s `autoCloseStream(false)` behavior: the caller retains and
    /// may continue using the stream after [`ExcelOutputStreamBuilder::do_write`].
    #[must_use]
    pub fn to_writer<W>(self, output: &mut W) -> ExcelOutputStreamBuilder<'_, T, W>
    where
        W: Write + Send,
    {
        ExcelOutputStreamBuilder {
            builder: self,
            output,
        }
    }

    /// Redirects this builder to a cloneable, explicitly closeable stream.
    ///
    /// This form supports both one-shot writes and stateful multi-batch writes,
    /// including Java-compatible `autoCloseStream` behavior.
    #[must_use]
    pub fn to_output_stream<W>(
        self,
        output: ExcelOutputStream<W>,
    ) -> ExcelOwnedOutputStreamBuilder<T, W>
    where
        W: Write + Send + 'static,
    {
        ExcelOwnedOutputStreamBuilder {
            builder: self,
            output,
        }
    }

    /// Enables or disables closing an owned output stream during finish.
    #[must_use]
    pub const fn auto_close_stream(mut self, enabled: bool) -> Self {
        self.options.auto_close_stream = enabled;
        self
    }

    /// Controls whether accumulated rows are emitted by `finish_on_exception`.
    #[must_use]
    pub const fn write_excel_on_exception(mut self, enabled: bool) -> Self {
        self.options.write_excel_on_exception = enabled;
        self
    }

    /// Builds a stateful writer for multiple `.write(rows, &sheet)` calls.
    #[must_use]
    pub fn build(self) -> ExcelWriter {
        ExcelWriter::with_handlers_and_options(self.path, self.handlers, self.options)
    }

    /// Selects constant-memory output.
    #[must_use]
    pub const fn constant_memory(mut self, enabled: bool) -> Self {
        self.options.constant_memory = enabled;
        self
    }

    /// Enables SXSSF-style compressed / disk-spill temporary files for bulk writes.
    ///
    /// Java mapping: `SXSSFWorkbook.setCompressTempFiles(true)` (commonly set in
    /// `WorkbookWriteHandler.afterWorkbookCreate`). Forces constant-memory row
    /// spill so large multi-batch writes do not keep the full sheet in RAM.
    ///
    /// See [`WriteOptions::compress_temp_files`] for the POI vs `rust_xlsxwriter`
    /// gzip difference.
    #[must_use]
    pub const fn compress_temp_files(mut self, enabled: bool) -> Self {
        self.options.compress_temp_files = enabled;
        if enabled {
            self.options.constant_memory = true;
        }
        self
    }

    /// Writes any owned row iterator.
    ///
    /// When [`Self::with_template`] is set, rows are appended onto the template
    /// workbook (Java `withTemplate(...).sheet().doWrite(data)`).
    ///
    /// # Errors
    ///
    /// Returns a conversion, worksheet-configuration, XLSX-format, template, or I/O error.
    pub fn do_write<I>(mut self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        let has_template =
            self.options.template_file.is_some() || self.options.template_bytes.is_some();
        if is_csv_path(&self.path) {
            if has_template {
                return Err(ExcelError::Unsupported(
                    "csv cannot use template.".to_owned(),
                ));
            }
            write_csv_with_handlers::<T, I>(
                Path::new(&self.path),
                &self.options,
                rows,
                &mut self.handlers,
            )
        } else if is_xls_path(&self.path) {
            // Java: EasyExcel.write(...).excelType(ExcelTypeEnum.XLS).sheet().doWrite(...)
            // Minimal BIFF8; with_template uses value-preserving rewrite (see biff8::template).
            write_xls_with_handlers::<T, I>(
                Path::new(&self.path),
                &self.options,
                rows,
                &mut self.handlers,
            )
        } else {
            write_xlsx_with_handlers::<T, I>(
                Path::new(&self.path),
                &self.options,
                rows,
                &mut self.handlers,
            )
        }
    }

    /// Alias emphasizing that the input is consumed incrementally.
    ///
    /// # Errors
    ///
    /// Returns a conversion, worksheet-configuration, XLSX-format, or I/O error.
    pub fn do_write_iter<I>(self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        self.do_write(rows)
    }

    /// Fills scalar `{key}` placeholders through [`ExcelBuilderImpl::fill`].
    ///
    /// Mirrors Java `EasyExcel.write(file).withTemplate(template).sheet().doFill(data)`.
    ///
    /// # Errors
    ///
    /// Returns template, fill, CSV/XLS unsupported, or output errors.
    pub fn do_fill(self, data: &TemplateData) -> Result<()> {
        let sheet = WriteSheet::<DynamicRow>::from_options(self.options.clone());
        do_fill_template(self.build(), data, &sheet)
    }
}

/// Caller-owned XLSX output stream builder.
pub struct ExcelOutputStreamBuilder<'a, T, W> {
    builder: ExcelWriterBuilder<T>,
    output: &'a mut W,
}

impl<T, W> ExcelOutputStreamBuilder<'_, T, W>
where
    T: ExcelRow,
    W: Write + Send,
{
    /// Writes a complete OOXML package to the borrowed stream and flushes it.
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, workbook, encryption, or stream I/O error.
    pub fn do_write<I>(mut self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        let has_template = self.builder.options.template_file.is_some()
            || self.builder.options.template_bytes.is_some();
        if is_csv_path(&self.builder.path) {
            if has_template {
                return Err(ExcelError::Unsupported(
                    "csv cannot use template.".to_owned(),
                ));
            }
            let bytes = write_csv_to_buffer::<T, I>(
                &self.builder.path,
                &self.builder.options,
                rows,
                &mut self.builder.handlers,
            )?;
            self.output.write_all(&bytes)?;
            self.output.flush()?;
            return Ok(());
        }
        if is_xls_path(&self.builder.path) {
            // Java stream write with ExcelTypeEnum.XLS — BIFF8 (+ optional template).
            return write_xls_to_writer::<T, I, _>(
                &self.builder.path,
                &mut *self.output,
                &self.builder.options,
                rows,
                &mut self.builder.handlers,
            );
        }
        write_xlsx_to_writer::<T, I, _>(
            &self.builder.path,
            self.output,
            &self.builder.options,
            rows,
            &mut self.builder.handlers,
        )
    }
}

/// Owned, cloneable output-stream builder for one-shot or stateful writes.
pub struct ExcelOwnedOutputStreamBuilder<T, W> {
    builder: ExcelWriterBuilder<T>,
    output: ExcelOutputStream<W>,
}

impl<T, W> ExcelOwnedOutputStreamBuilder<T, W>
where
    T: ExcelRow,
    W: Write + Send + 'static,
{
    /// Builds a stateful writer for repeated `write` calls.
    #[must_use]
    pub fn build(self) -> ExcelWriter {
        ExcelWriter::with_output_stream(
            self.builder.path,
            self.output,
            self.builder.handlers,
            self.builder.options,
        )
    }

    /// Writes one batch and completes the output-stream lifecycle.
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, workbook, close, or stream I/O error.
    pub fn do_write<I>(self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        let sheet = WriteSheet::from_options(self.builder.options.clone());
        let mut writer = self.build();
        if let Err(error) = writer.write(rows, &sheet) {
            writer.finish_on_exception()?;
            return Err(error);
        }
        writer.finish()
    }
}

fn is_csv_path(path: &Path) -> bool {
    path.extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("csv"))
}

fn is_xls_path(path: &Path) -> bool {
    path.extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
}

#[cfg(test)]
mod tests;
