//! Public facade for typed, event-driven Excel reading and writing.

use std::marker::PhantomData;
use std::path::{Path, PathBuf};

pub use easyexcel_core::*;
pub use easyexcel_derive::ExcelRow;
pub use easyexcel_reader::ExcelLocale;
use easyexcel_reader::{
    ReadOptions, ScientificFormatMode, SheetSelector, read_csv, read_xls, read_xlsx,
};
pub use easyexcel_template::{
    FillConfig, FillDirection, FillWrapper, TemplateData, fill_xlsx_template,
    fill_xlsx_template_list,
};
pub use easyexcel_writer::{
    CellStyle, ExcelWriter, HorizontalAlignment, LoopMergeStrategy, MergeRange, VerticalAlignment,
    WriteOptions, WriteSheet, write_csv_to_writer,
};
use easyexcel_writer::{write_csv_with_handlers, write_xlsx_with_handlers};

/// Static factory matching Java `EasyExcel`'s entry point.
pub struct EasyExcel;

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

    /// Fills scalar `{key}` placeholders in an existing XLSX template.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error.
    pub fn fill_template(
        template: impl AsRef<Path>,
        output: impl AsRef<Path>,
        data: &TemplateData,
    ) -> Result<()> {
        fill_xlsx_template(template.as_ref(), output.as_ref(), data)
    }

    /// Expands a collection in an existing XLSX template.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error.
    pub fn fill_template_list(
        template: impl AsRef<Path>,
        output: impl AsRef<Path>,
        data: &FillWrapper,
        config: FillConfig,
    ) -> Result<()> {
        fill_xlsx_template_list(template.as_ref(), output.as_ref(), data, config)
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

    /// Writes any owned row iterator.
    ///
    /// # Errors
    ///
    /// Returns a conversion, worksheet-configuration, XLSX-format, or I/O error.
    pub fn do_write<I>(mut self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        if is_csv_path(&self.path) {
            write_csv_with_handlers::<T, I>(
                Path::new(&self.path),
                &self.options,
                rows,
                &mut self.handlers,
            )
        } else if is_xls_path(&self.path) {
            Err(ExcelError::Unsupported(
                "legacy XLS writing is not supported".to_owned(),
            ))
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
