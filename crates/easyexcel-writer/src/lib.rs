//! XLSX writer backed by `rust_xlsxwriter`.

use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Read, Seek, Write};
use std::marker::PhantomData;
use std::path::{Path, PathBuf};

use bigdecimal::ToPrimitive;
use easyexcel_core::{
    AnchorType, CellValue, Converter, ConverterRegistry, CsvCharset, ExcelBorderStyle,
    ExcelCellStyle, ExcelColor, ExcelColumn, ExcelDataFormat, ExcelError, ExcelFillPattern,
    ExcelFontScript, ExcelFontStyle, ExcelHorizontalAlignment, ExcelRow, ExcelUnderline,
    ExcelVerticalAlignment, ExcelWriteMetadata, ImageData, Result, WriteCellContext, WriteHandler,
    WriteRowContext, WriteSheetContext, WriteWorkbookContext,
};
use encoding_rs::{CoderResult, Encoding, UTF_8, UTF_16BE, UTF_16LE};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, FormatPattern, FormatScript, FormatUnderline, Image,
    Note, ObjectMovement, Workbook, Worksheet,
};

/// Horizontal cell alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalAlignment {
    /// Excel's type-dependent default.
    General,
    /// Left aligned.
    Left,
    /// Centered.
    Center,
    /// Right aligned.
    Right,
    /// Repeats content across the cell.
    Fill,
    /// Justified.
    Justify,
    /// Centered across adjacent cells.
    CenterAcross,
}

/// Vertical cell alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    /// Top aligned.
    Top,
    /// Vertically centered.
    Center,
    /// Bottom aligned.
    Bottom,
    /// Vertically justified.
    Justify,
    /// Vertically distributed.
    Distributed,
}

/// Backend-neutral write style for headers or content rows.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellStyle {
    /// Bold font.
    pub bold: bool,
    /// Italic font.
    pub italic: bool,
    /// RGB font color, for example `0xFF0000`.
    pub font_color: Option<u32>,
    /// Solid RGB background color.
    pub background_color: Option<u32>,
    /// Horizontal alignment.
    pub horizontal_alignment: Option<HorizontalAlignment>,
    /// Vertical alignment.
    pub vertical_alignment: Option<VerticalAlignment>,
    /// Wrap cell text.
    pub wrap_text: bool,
    /// Excel number format string.
    pub number_format: Option<String>,
}

impl CellStyle {
    /// Creates an empty style.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            bold: false,
            italic: false,
            font_color: None,
            background_color: None,
            horizontal_alignment: None,
            vertical_alignment: None,
            wrap_text: false,
            number_format: None,
        }
    }

    /// Sets bold font rendering.
    #[must_use]
    pub const fn bold(mut self, enabled: bool) -> Self {
        self.bold = enabled;
        self
    }

    /// Sets italic font rendering.
    #[must_use]
    pub const fn italic(mut self, enabled: bool) -> Self {
        self.italic = enabled;
        self
    }

    /// Sets the RGB font color.
    #[must_use]
    pub const fn font_color(mut self, color: u32) -> Self {
        self.font_color = Some(color);
        self
    }

    /// Sets a solid RGB background color.
    #[must_use]
    pub const fn background_color(mut self, color: u32) -> Self {
        self.background_color = Some(color);
        self
    }

    /// Sets horizontal alignment.
    #[must_use]
    pub const fn horizontal_alignment(mut self, alignment: HorizontalAlignment) -> Self {
        self.horizontal_alignment = Some(alignment);
        self
    }

    /// Sets vertical alignment.
    #[must_use]
    pub const fn vertical_alignment(mut self, alignment: VerticalAlignment) -> Self {
        self.vertical_alignment = Some(alignment);
        self
    }

    /// Enables or disables text wrapping.
    #[must_use]
    pub const fn wrap_text(mut self, enabled: bool) -> Self {
        self.wrap_text = enabled;
        self
    }

    /// Sets an Excel number format string.
    #[must_use]
    pub fn number_format(mut self, format: impl Into<String>) -> Self {
        self.number_format = Some(format.into());
        self
    }
}

/// One absolute merged-cell range using zero-based inclusive coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MergeRange {
    /// First row.
    pub first_row: u32,
    /// Last row.
    pub last_row: u32,
    /// First column.
    pub first_column: u16,
    /// Last column.
    pub last_column: u16,
}

impl MergeRange {
    /// Creates an absolute merge range.
    #[must_use]
    pub const fn new(first_row: u32, last_row: u32, first_column: u16, last_column: u16) -> Self {
        Self {
            first_row,
            last_row,
            first_column,
            last_column,
        }
    }
}

/// Repeating merge strategy equivalent to Java `LoopMergeStrategy`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LoopMergeStrategy {
    each_rows: u32,
    column_extend: u16,
    column_index: u16,
}

impl LoopMergeStrategy {
    /// Creates a repeating merge strategy.
    ///
    /// # Errors
    ///
    /// Returns a format error when dimensions are zero or no range would be merged.
    pub fn new(each_rows: u32, column_extend: u16, column_index: u16) -> Result<Self> {
        if each_rows == 0 || column_extend == 0 {
            return Err(ExcelError::Format(
                "loop merge row and column spans must be greater than zero".to_owned(),
            ));
        }
        if each_rows == 1 && column_extend == 1 {
            return Err(ExcelError::Format(
                "loop merge must span multiple rows or columns".to_owned(),
            ));
        }
        Ok(Self {
            each_rows,
            column_extend,
            column_index,
        })
    }

    /// Number of rows in each merge group.
    #[must_use]
    pub const fn each_rows(self) -> u32 {
        self.each_rows
    }

    /// Number of columns in each merge group.
    #[must_use]
    pub const fn column_extend(self) -> u16 {
        self.column_extend
    }

    /// Zero-based first column.
    #[must_use]
    pub const fn column_index(self) -> u16 {
        self.column_index
    }
}

/// XLSX write configuration.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(clippy::struct_excessive_bools)]
pub struct WriteOptions {
    /// Worksheet name.
    pub sheet_name: String,
    /// Optional logical worksheet number, starting from zero.
    pub sheet_index: Option<usize>,
    /// Whether to use a one-row constant-memory worksheet.
    pub constant_memory: bool,
    /// Whether column headers are written.
    pub need_head: bool,
    /// Whether header rows are frozen.
    pub freeze_head: bool,
    /// Explicit freeze pane position as `(row, column)`.
    pub freeze_panes: Option<(u32, u16)>,
    /// Physical column indexes to include.
    pub include_column_indexes: Option<Vec<usize>>,
    /// Rust field names to include.
    pub include_column_field_names: Option<Vec<String>>,
    /// Physical column indexes to exclude.
    pub exclude_column_indexes: Vec<usize>,
    /// Rust field names to exclude.
    pub exclude_column_field_names: Vec<String>,
    /// Whether included columns follow the order of the include list.
    pub order_by_include_column: bool,
    /// Absolute ranges merged before row data is written.
    pub merge_ranges: Vec<MergeRange>,
    /// Whether used columns are auto-fitted after writing.
    pub auto_width: bool,
    /// Explicit column widths in Excel character units.
    pub column_widths: Vec<(u16, u16)>,
    /// Style applied to header cells.
    pub head_style: CellStyle,
    /// Content styles cycled by relative data-row index.
    pub content_styles: Vec<CellStyle>,
    /// Repeating merge strategies applied to data rows.
    pub loop_merges: Vec<LoopMergeStrategy>,
    /// Optional dynamic multi-level head paths, one path per selected column.
    pub dynamic_head: Option<Vec<Vec<String>>>,
    /// Password used for ECMA-376 Agile Encryption of XLSX output.
    pub password: Option<String>,
    /// Character encoding used for CSV output.
    pub charset: CsvCharset,
    /// Whether CSV output starts with the encoding's byte-order mark.
    pub with_bom: bool,
    /// Java-style globally registered converters.
    pub converters: ConverterRegistry,
}

impl Default for WriteOptions {
    fn default() -> Self {
        Self {
            sheet_name: "Sheet1".to_owned(),
            sheet_index: None,
            constant_memory: false,
            need_head: true,
            freeze_head: false,
            freeze_panes: None,
            include_column_indexes: None,
            include_column_field_names: None,
            exclude_column_indexes: Vec::new(),
            exclude_column_field_names: Vec::new(),
            order_by_include_column: false,
            merge_ranges: Vec::new(),
            auto_width: false,
            column_widths: Vec::new(),
            head_style: CellStyle::new().bold(true),
            content_styles: Vec::new(),
            loop_merges: Vec::new(),
            dynamic_head: None,
            password: None,
            charset: CsvCharset::default(),
            with_bom: true,
            converters: ConverterRegistry::default(),
        }
    }
}

/// Typed worksheet metadata used by [`ExcelWriter`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheet<T> {
    options: WriteOptions,
    marker: PhantomData<T>,
}

impl<T> WriteSheet<T> {
    /// Creates worksheet metadata with the supplied name.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            options: WriteOptions {
                sheet_name: name.into(),
                ..WriteOptions::default()
            },
            marker: PhantomData,
        }
    }

    /// Creates worksheet metadata identified by a Java-style zero-based sheet number.
    #[must_use]
    pub fn new_index(index: usize) -> Self {
        Self {
            options: WriteOptions {
                sheet_name: index.to_string(),
                sheet_index: Some(index),
                ..WriteOptions::default()
            },
            marker: PhantomData,
        }
    }

    /// Returns the effective write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }

    /// Registers a sheet-level converter that overrides a workbook registration.
    #[must_use]
    pub fn register_converter<V, C>(mut self, converter: C) -> Self
    where
        V: 'static,
        C: Converter<V> + Send + Sync + 'static,
    {
        self.options.converters.register::<V, C>(converter);
        self
    }

    /// Adds a Java-style zero-based logical sheet number to this worksheet.
    #[must_use]
    pub const fn sheet_index(mut self, index: usize) -> Self {
        self.options.sheet_index = Some(index);
        self
    }

    /// Enables or disables headers for this sheet.
    #[must_use]
    pub const fn need_head(mut self, enabled: bool) -> Self {
        self.options.need_head = enabled;
        self
    }

    /// Enables or disables constant-memory output for this sheet.
    #[must_use]
    pub const fn constant_memory(mut self, enabled: bool) -> Self {
        self.options.constant_memory = enabled;
        self
    }

    /// Freezes the header row for this sheet.
    #[must_use]
    pub const fn freeze_head(mut self, enabled: bool) -> Self {
        self.options.freeze_head = enabled;
        self
    }

    /// Adds an absolute merged-cell range.
    #[must_use]
    pub fn merge_cells(mut self, range: MergeRange) -> Self {
        self.options.merge_ranges.push(range);
        self
    }

    /// Enables or disables automatic width calculation.
    #[must_use]
    pub const fn auto_width(mut self, enabled: bool) -> Self {
        self.options.auto_width = enabled;
        self
    }

    /// Sets an explicit width for a zero-based physical column.
    #[must_use]
    pub fn column_width(mut self, column: u16, width: u16) -> Self {
        self.options.column_widths.push((column, width));
        self
    }

    /// Replaces the header style.
    #[must_use]
    pub fn head_style(mut self, style: CellStyle) -> Self {
        self.options.head_style = style;
        self
    }

    /// Uses one style for every content row.
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

    /// Registers a repeating data-row merge strategy.
    #[must_use]
    pub fn loop_merge(mut self, strategy: LoopMergeStrategy) -> Self {
        self.options.loop_merges.push(strategy);
        self
    }

    /// Replaces the derived headers with dynamic multi-level head paths.
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
}

#[derive(Clone)]
struct StatefulSheetState {
    schema: &'static [ExcelColumn],
    metadata: ExcelWriteMetadata,
    options: WriteOptions,
    next_row: u32,
    next_data_index: usize,
}

/// Stateful XLSX or single-sheet CSV writer matching Java `ExcelWriter`'s lifecycle.
pub struct ExcelWriter {
    path: PathBuf,
    workbook: Workbook,
    handlers: Vec<Box<dyn WriteHandler>>,
    sheets: HashMap<String, StatefulSheetState>,
    sheet_indexes: HashMap<usize, String>,
    csv_writer: Option<csv::Writer<CsvEncodingWriter>>,
    csv_charset: CsvCharset,
    csv_with_bom: bool,
    started: bool,
    finished: bool,
    password: Option<String>,
    converters: ConverterRegistry,
}

impl ExcelWriter {
    /// Creates a multi-sheet writer without handlers.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        Self::with_handlers(path, Vec::new())
    }

    /// Creates a multi-sheet writer with owned lifecycle handlers.
    #[must_use]
    pub fn with_handlers(path: impl Into<PathBuf>, handlers: Vec<Box<dyn WriteHandler>>) -> Self {
        Self::with_handlers_and_password(path, handlers, None)
    }

    /// Creates a multi-sheet writer with handlers and optional XLSX encryption.
    #[must_use]
    pub fn with_handlers_and_password(
        path: impl Into<PathBuf>,
        handlers: Vec<Box<dyn WriteHandler>>,
        password: Option<String>,
    ) -> Self {
        Self::with_handlers_and_options(
            path,
            handlers,
            WriteOptions {
                password,
                ..WriteOptions::default()
            },
        )
    }

    /// Creates a stateful writer with workbook-level builder options.
    #[must_use]
    pub fn with_handlers_and_options(
        path: impl Into<PathBuf>,
        handlers: Vec<Box<dyn WriteHandler>>,
        options: WriteOptions,
    ) -> Self {
        Self {
            path: path.into(),
            workbook: Workbook::new(),
            handlers,
            sheets: HashMap::new(),
            sheet_indexes: HashMap::new(),
            csv_writer: None,
            csv_charset: options.charset,
            csv_with_bom: options.with_bom,
            started: false,
            finished: false,
            password: options.password,
            converters: options.converters,
        }
    }

    /// Writes a batch to a worksheet, appending when the sheet was used before.
    ///
    /// XLSX permits multiple sheets. CSV permits repeated writes to one logical
    /// sheet, matching Java `EasyExcel`'s stateful writer.
    ///
    /// # Errors
    ///
    /// Returns an error when the writer is finished, a handler fails, or data cannot be written.
    pub fn write<T, I>(&mut self, rows: I, sheet: &WriteSheet<T>) -> Result<&mut Self>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        if self.finished {
            return Err(ExcelError::Unsupported(
                "writer already finished".to_owned(),
            ));
        }
        self.start()?;
        if is_csv_path(&self.path) {
            self.write_csv_batch::<T, I>(rows, sheet)?;
        } else {
            self.write_xlsx_batch::<T, I>(rows, sheet)?;
        }
        debug_assert!(self.resolve_sheet_name(sheet.options()).is_some());
        Ok(self)
    }

    /// Saves and closes the writer. Repeated calls are no-ops.
    ///
    /// # Errors
    ///
    /// Returns an output or handler error.
    pub fn finish(&mut self) -> Result<()> {
        if self.finished {
            return Ok(());
        }
        self.start()?;
        if is_csv_path(&self.path) {
            finish_stateful_csv_writer(&mut self.csv_writer)?;
        } else {
            save_workbook(&mut self.workbook, &self.path, self.password.as_deref())?;
        }
        let context = WriteWorkbookContext::new(&self.path);
        after_workbook(&mut self.handlers, &context)?;
        self.finished = true;
        Ok(())
    }

    /// Returns whether [`Self::finish`] completed successfully.
    #[must_use]
    pub const fn is_finished(&self) -> bool {
        self.finished
    }

    fn start(&mut self) -> Result<()> {
        if self.started {
            return Ok(());
        }
        validate_stateful_backend(&self.path, self.password.as_deref())?;
        sort_handlers(&mut self.handlers);
        let context = WriteWorkbookContext::new(&self.path);
        before_workbook(&mut self.handlers, &context)?;
        if is_csv_path(&self.path) {
            self.csv_writer = Some(create_stateful_csv_writer(
                &self.path,
                &self.csv_charset,
                self.csv_with_bom,
            )?);
        }
        self.started = true;
        Ok(())
    }

    fn write_xlsx_batch<T, I>(&mut self, rows: I, sheet: &WriteSheet<T>) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let requested_name = sheet.options().sheet_name.clone();
        if let Some(sheet_name) = self.resolve_sheet_name(sheet.options()) {
            let state = self
                .sheets
                .get(&sheet_name)
                .cloned()
                .expect("resolved worksheet must exist");
            validate_stateful_schema(&sheet_name, &state, T::schema())?;
            let worksheet = self
                .workbook
                .worksheet_from_name(&sheet_name)
                .map_err(format_error)?;
            let progress = append_rows_to_worksheet::<T, I>(
                worksheet,
                &state.options,
                rows,
                &mut self.handlers,
                WriteProgress {
                    next_row: state.next_row,
                    next_data_index: state.next_data_index,
                },
                false,
                &state.metadata,
            )?;
            if state.options.auto_width {
                worksheet.autofit();
            }
            let current = self
                .sheets
                .get_mut(&sheet_name)
                .expect("stateful worksheet must exist");
            current.next_row = progress.next_row;
            current.next_data_index = progress.next_data_index;
            return Ok(());
        }

        let mut options = sheet.options().clone();
        options.converters = self.converters.merged_with(&options.converters);
        let progress = write_sheet_to_workbook::<T, I>(
            &mut self.workbook,
            &options,
            rows,
            &mut self.handlers,
        )?;
        self.sheets.insert(
            requested_name.clone(),
            StatefulSheetState {
                schema: T::schema(),
                metadata: *T::write_metadata(),
                options,
                next_row: progress.next_row,
                next_data_index: progress.next_data_index,
            },
        );
        self.remember_sheet_index(sheet.options().sheet_index, &requested_name);
        Ok(())
    }

    fn write_csv_batch<T, I>(&mut self, rows: I, sheet: &WriteSheet<T>) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let requested_name = sheet.options().sheet_name.clone();
        let existing_name = self.resolve_sheet_name(sheet.options());
        if existing_name.is_none() && !self.sheets.is_empty() {
            return Err(ExcelError::Unsupported(
                "CSV supports only one worksheet".to_owned(),
            ));
        }
        let sheet_name = existing_name.unwrap_or(requested_name);

        let (state, is_new) = if let Some(state) = self.sheets.get(&sheet_name).cloned() {
            validate_stateful_schema(&sheet_name, &state, T::schema())?;
            (state, false)
        } else {
            let mut options = sheet.options().clone();
            options.charset = self.csv_charset.clone();
            options.with_bom = self.csv_with_bom;
            options.converters = self.converters.merged_with(&options.converters);
            (
                StatefulSheetState {
                    schema: T::schema(),
                    metadata: *T::write_metadata(),
                    options,
                    next_row: 0,
                    next_data_index: 0,
                },
                true,
            )
        };

        let sheet_context = WriteSheetContext::new(&sheet_name);
        if is_new {
            before_sheet(&mut self.handlers, &sheet_context)?;
        }
        let writer = self
            .csv_writer
            .as_mut()
            .expect("stateful CSV writer must be initialized");
        let progress = append_csv_rows::<T, I>(
            writer,
            &state.options,
            rows,
            &mut self.handlers,
            state.next_row,
            state.next_data_index,
            is_new,
        )?;
        if is_new {
            after_sheet(&mut self.handlers, &sheet_context)?;
        }
        self.sheets.insert(
            sheet_name.clone(),
            StatefulSheetState {
                next_row: progress.next_row,
                next_data_index: progress.next_data_index,
                ..state
            },
        );
        if is_new {
            self.remember_sheet_index(sheet.options().sheet_index, &sheet_name);
        }
        Ok(())
    }

    fn resolve_sheet_name(&self, options: &WriteOptions) -> Option<String> {
        options
            .sheet_index
            .and_then(|index| self.sheet_indexes.get(&index).cloned())
            .or_else(|| {
                self.sheets
                    .contains_key(&options.sheet_name)
                    .then(|| options.sheet_name.clone())
            })
    }

    fn remember_sheet_index(&mut self, index: Option<usize>, sheet_name: &str) {
        if let Some(index) = index {
            self.sheet_indexes.insert(index, sheet_name.to_owned());
        }
    }
}

fn validate_stateful_backend(path: &Path, password: Option<&str>) -> Result<()> {
    match path.extension().and_then(std::ffi::OsStr::to_str) {
        Some(extension) if extension.eq_ignore_ascii_case("csv") && password.is_some() => Err(
            ExcelError::Unsupported("password protection is not supported for CSV".to_owned()),
        ),
        Some(extension) if extension.eq_ignore_ascii_case("xls") => Err(ExcelError::Unsupported(
            "legacy XLS writing is not supported".to_owned(),
        )),
        _ => Ok(()),
    }
}

fn is_csv_path(path: &Path) -> bool {
    path.extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("csv"))
}

fn validate_stateful_schema(
    sheet_name: &str,
    state: &StatefulSheetState,
    schema: &'static [ExcelColumn],
) -> Result<()> {
    if state.schema == schema {
        Ok(())
    } else {
        Err(ExcelError::Format(format!(
            "worksheet schema changed between writes: {sheet_name}"
        )))
    }
}

/// Writes typed rows to a new XLSX file.
///
/// # Errors
///
/// Returns a conversion, worksheet-configuration, XLSX-format, or I/O error.
pub fn write_xlsx<T, I>(path: &Path, options: &WriteOptions, rows: I) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    write_xlsx_with_handlers(path, options, rows, &mut [])
}

/// Writes typed rows while invoking ordered write handlers.
///
/// # Errors
///
/// Returns a conversion, handler, worksheet-configuration, XLSX-format, or I/O error.
pub fn write_xlsx_with_handlers<T, I>(
    path: &Path,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    sort_handlers(handlers);
    let workbook_context = WriteWorkbookContext::new(path);
    before_workbook(handlers, &workbook_context)?;

    let mut workbook = Workbook::new();
    write_sheet_to_workbook::<T, I>(&mut workbook, options, rows, handlers)?;
    save_workbook(&mut workbook, path, options.password.as_deref())?;
    after_workbook(handlers, &workbook_context)?;
    Ok(())
}

/// Writes typed rows to CSV while preserving the `EasyExcel` handler lifecycle.
///
/// UTF-8 BOM output matches Java `EasyExcel`'s default CSV behavior.
///
/// # Errors
///
/// Returns a conversion, handler, CSV-format, or I/O error.
pub fn write_csv_with_handlers<T, I>(
    path: &Path,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    validate_csv_options(options)?;
    let file = File::create(path)?;
    write_csv_to::<T, I>(path, Box::new(file), options, rows, handlers)
}

/// Writes typed CSV rows to an owned byte stream.
///
/// `logical_path` is used by write-handler contexts and does not need to exist
/// on the filesystem. This is the Rust equivalent of Java `EasyExcel`'s
/// `OutputStream` CSV entry point.
///
/// # Errors
///
/// Returns a conversion, handler, CSV-format, charset, or stream I/O error.
pub fn write_csv_to_writer<T, I, W>(
    logical_path: &Path,
    output: W,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
    W: Write + 'static,
{
    validate_csv_options(options)?;
    write_csv_to::<T, I>(logical_path, Box::new(output), options, rows, handlers)
}

fn write_csv_to<T, I>(
    path: &Path,
    output: Box<dyn Write>,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let columns = selected_columns(T::schema(), options);
    let mut rows = rows
        .into_iter()
        .map(|row| row.to_row_with_converters(&options.converters));
    write_csv_records(
        path,
        output,
        options,
        &columns,
        T::schema().is_empty(),
        &mut rows,
        handlers,
    )
}

fn write_csv_records(
    path: &Path,
    output: Box<dyn Write>,
    options: &WriteOptions,
    columns: &[(usize, usize, &'static ExcelColumn)],
    schema_is_empty: bool,
    rows: &mut dyn Iterator<Item = Result<Vec<CellValue>>>,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    csv_encoding(&options.charset)?;
    sort_handlers(handlers);
    let workbook_context = WriteWorkbookContext::new(path);
    before_workbook(handlers, &workbook_context)?;
    let sheet_context = WriteSheetContext::new(&options.sheet_name);
    before_sheet(handlers, &sheet_context)?;

    let mut writer = create_csv_record_writer(output, &options.charset, options.with_bom)?;
    append_csv_records(
        &mut writer,
        options,
        columns,
        schema_is_empty,
        rows,
        handlers,
        0,
        0,
        true,
    )?;
    finish_csv_record_writer(writer)?;
    after_sheet(handlers, &sheet_context)?;
    after_workbook(handlers, &workbook_context)
}

#[allow(clippy::too_many_arguments)]
fn append_csv_records(
    writer: &mut csv::Writer<CsvEncodingWriter>,
    options: &WriteOptions,
    columns: &[(usize, usize, &'static ExcelColumn)],
    schema_is_empty: bool,
    rows: &mut dyn Iterator<Item = Result<Vec<CellValue>>>,
    handlers: &mut [Box<dyn WriteHandler>],
    mut row_index: u32,
    mut data_index: usize,
    write_head: bool,
) -> Result<WriteProgress> {
    let head_rows = head_rows_for_schema_state(schema_is_empty, options)?;
    if write_head && head_rows > 0 {
        if let Some(head) = &options.dynamic_head {
            if head.len() != columns.len() {
                return Err(ExcelError::Format(format!(
                    "dynamic head column count {} does not match selected column count {}",
                    head.len(),
                    columns.len()
                )));
            }
            for level in 0..head_rows {
                #[allow(clippy::cast_possible_truncation)]
                let level = level as usize;
                let labels = head
                    .iter()
                    .map(|path| path.get(level).cloned().unwrap_or_default())
                    .collect::<Vec<_>>();
                let record =
                    csv_header_record(row_index, columns, &labels, &options.sheet_name, handlers)?;
                writer.write_record(record).map_err(format_error)?;
                row_index += 1;
            }
        } else {
            let labels = columns
                .iter()
                .map(|(_, _, column)| column.name.to_owned())
                .collect::<Vec<_>>();
            let record =
                csv_header_record(row_index, columns, &labels, &options.sheet_name, handlers)?;
            writer.write_record(record).map_err(format_error)?;
            row_index = 1;
        }
    }
    for cells in rows {
        let cells = cells?;
        let dynamic_columns = dynamic_columns_for_row(schema_is_empty, cells.len(), options);
        let row_columns = dynamic_columns.as_deref().unwrap_or(columns);
        let record = csv_data_record(
            row_index,
            row_columns,
            &cells,
            &options.sheet_name,
            handlers,
        )?;
        writer.write_record(record).map_err(format_error)?;
        row_index += 1;
        data_index += 1;
    }
    Ok(WriteProgress {
        next_row: row_index,
        next_data_index: data_index,
    })
}

fn append_csv_rows<T, I>(
    writer: &mut csv::Writer<CsvEncodingWriter>,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
    row_index: u32,
    data_index: usize,
    write_head: bool,
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let columns = selected_columns(T::schema(), options);
    let mut rows = rows
        .into_iter()
        .map(|row| row.to_row_with_converters(&options.converters));
    append_csv_records(
        writer,
        options,
        &columns,
        T::schema().is_empty(),
        &mut rows,
        handlers,
        row_index,
        data_index,
        write_head,
    )
}

fn create_csv_record_writer(
    mut output: Box<dyn Write>,
    charset: &CsvCharset,
    with_bom: bool,
) -> Result<csv::Writer<CsvEncodingWriter>> {
    let encoding = csv_encoding(charset)?;
    if with_bom {
        output.write_all(csv_bom(encoding))?;
    }
    Ok(csv::WriterBuilder::new().from_writer(CsvEncodingWriter::new(output, encoding)))
}

fn create_stateful_csv_writer(
    path: &Path,
    charset: &CsvCharset,
    with_bom: bool,
) -> Result<csv::Writer<CsvEncodingWriter>> {
    create_csv_record_writer(Box::new(File::create(path)?), charset, with_bom)
}

fn finish_csv_record_writer(mut writer: csv::Writer<CsvEncodingWriter>) -> Result<()> {
    writer.flush()?;
    let mut output = writer.into_inner().map_err(format_error)?;
    output.finish()?;
    Ok(())
}

fn finish_stateful_csv_writer(writer: &mut Option<csv::Writer<CsvEncodingWriter>>) -> Result<()> {
    if let Some(writer) = writer.take() {
        finish_csv_record_writer(writer)?;
    }
    Ok(())
}

fn validate_csv_options(options: &WriteOptions) -> Result<()> {
    if options.password.is_some() {
        return Err(ExcelError::Unsupported(
            "password protection is not supported for CSV".to_owned(),
        ));
    }
    csv_encoding(&options.charset)?;
    Ok(())
}

fn csv_encoding(charset: &CsvCharset) -> Result<CsvEncoding> {
    let encoding = Encoding::for_label(charset.name().as_bytes()).ok_or_else(|| {
        ExcelError::Unsupported(format!("unsupported CSV charset: {}", charset.name()))
    })?;
    Ok(if encoding == UTF_16LE {
        CsvEncoding::Utf16Le
    } else if encoding == UTF_16BE {
        CsvEncoding::Utf16Be
    } else {
        CsvEncoding::Standard(encoding)
    })
}

fn csv_bom(encoding: CsvEncoding) -> &'static [u8] {
    match encoding {
        CsvEncoding::Standard(encoding) if encoding == UTF_8 => b"\xEF\xBB\xBF",
        CsvEncoding::Utf16Le => b"\xFF\xFE",
        CsvEncoding::Utf16Be => b"\xFE\xFF",
        CsvEncoding::Standard(_) => b"",
    }
}

#[derive(Clone, Copy)]
enum CsvEncoding {
    Standard(&'static Encoding),
    Utf16Le,
    Utf16Be,
}

enum CsvEncoder {
    Standard(encoding_rs::Encoder),
    Utf16Le,
    Utf16Be,
}

struct CsvEncodingWriter {
    output: Box<dyn Write>,
    encoder: CsvEncoder,
    pending_utf8: Vec<u8>,
}

impl CsvEncodingWriter {
    fn new(output: Box<dyn Write>, encoding: CsvEncoding) -> Self {
        Self {
            output,
            encoder: match encoding {
                CsvEncoding::Standard(encoding) => CsvEncoder::Standard(encoding.new_encoder()),
                CsvEncoding::Utf16Le => CsvEncoder::Utf16Le,
                CsvEncoding::Utf16Be => CsvEncoder::Utf16Be,
            },
            pending_utf8: Vec::new(),
        }
    }

    fn encode_text(&mut self, text: &str, last: bool) -> std::io::Result<()> {
        match &mut self.encoder {
            CsvEncoder::Standard(encoder) => {
                Self::encode_standard(&mut self.output, encoder, text, last)
            }
            CsvEncoder::Utf16Le => Self::encode_utf16(&mut self.output, text, u16::to_le_bytes),
            CsvEncoder::Utf16Be => Self::encode_utf16(&mut self.output, text, u16::to_be_bytes),
        }
    }

    fn encode_standard(
        output: &mut dyn Write,
        encoder: &mut encoding_rs::Encoder,
        mut text: &str,
        last: bool,
    ) -> std::io::Result<()> {
        loop {
            // Keep the transcoder chunk below csv's internal buffer so a
            // single upstream write can be continued without accumulating
            // the complete record in memory.
            let mut buffer = [0_u8; 4 * 1_024];
            let (result, read, written, _) = encoder.encode_from_utf8(text, &mut buffer, last);
            output.write_all(&buffer[..written])?;
            text = &text[read..];
            if result == CoderResult::InputEmpty {
                return Ok(());
            }
        }
    }

    fn encode_utf16(
        output: &mut dyn Write,
        text: &str,
        to_bytes: fn(u16) -> [u8; 2],
    ) -> std::io::Result<()> {
        let mut encoded = [0_u8; 8 * 1_024];
        let mut length = 0;
        for unit in text.encode_utf16() {
            if length == encoded.len() {
                output.write_all(&encoded)?;
                length = 0;
            }
            let bytes = to_bytes(unit);
            encoded[length] = bytes[0];
            encoded[length + 1] = bytes[1];
            length += 2;
        }
        output.write_all(&encoded[..length])
    }

    fn finish(&mut self) -> std::io::Result<()> {
        if !self.pending_utf8.is_empty() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                "CSV writer ended with incomplete UTF-8",
            ));
        }
        self.encode_text("", true)?;
        self.output.flush()
    }
}

impl Write for CsvEncodingWriter {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.pending_utf8.extend_from_slice(buffer);
        let valid_length = match std::str::from_utf8(&self.pending_utf8) {
            Ok(_) => self.pending_utf8.len(),
            Err(error) if error.error_len().is_none() => error.valid_up_to(),
            Err(error) => {
                return Err(std::io::Error::new(std::io::ErrorKind::InvalidData, error));
            }
        };
        if valid_length > 0 {
            let valid = self.pending_utf8.drain(..valid_length).collect::<Vec<_>>();
            let text = String::from_utf8_lossy(&valid);
            self.encode_text(text.as_ref(), false)?;
        }
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.output.flush()
    }
}

fn save_workbook(workbook: &mut Workbook, path: &Path, password: Option<&str>) -> Result<()> {
    let Some(password) = password else {
        return workbook.save(path).map_err(format_error);
    };
    let mut file = OpenOptions::new()
        .read(true)
        .write(true)
        .create(true)
        .truncate(true)
        .open(path)?;
    save_encrypted_workbook_to(workbook, password, &mut file)
}

trait ReadWriteSeek: Read + Write + Seek {}

impl<T> ReadWriteSeek for T where T: Read + Write + Seek {}

fn save_encrypted_workbook_to(
    workbook: &mut Workbook,
    password: &str,
    file: &mut dyn ReadWriteSeek,
) -> Result<()> {
    let mut random = rand::rng();
    Ecma376AgileWriter::create(&mut random, password, file)
        .map_err(ExcelError::from)
        .and_then(|mut writer| {
            workbook
                .save_to_buffer()
                .map_err(format_error)
                .and_then(|plaintext| {
                    // The encryption crate writes plaintext only to its in-memory cursor; its
                    // `Write` implementation cannot reach the fallible output at this stage.
                    let _ = writer.write_all(&plaintext);
                    writer.finalize().map_err(ExcelError::from)
                })
        })
}

fn csv_header_record(
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    labels: &[String],
    sheet_name: &str,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<Vec<String>> {
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: true,
    };
    before_csv_row(handlers, &row_context)?;
    let mut record = csv_record(columns);
    for ((physical_index, _, column), label) in columns.iter().zip(labels) {
        let column_index = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index,
            field: (!column.field.is_empty()).then_some(column.field),
            is_head: true,
            value: CellValue::String(label.clone()),
            skip: false,
        };
        before_csv_cell(handlers, &mut context)?;
        if !context.skip {
            record[*physical_index] = context.value.as_text();
        }
        after_csv_cell(handlers, &context)?;
    }
    after_csv_row(handlers, &row_context)?;
    Ok(record)
}

fn csv_data_record(
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
    sheet_name: &str,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<Vec<String>> {
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: false,
    };
    before_csv_row(handlers, &row_context)?;
    let mut record = csv_record(columns);
    for (physical_index, schema_index, metadata) in columns {
        let column_index = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index,
            field: (!metadata.field.is_empty()).then_some(metadata.field),
            is_head: false,
            value: cells
                .get(*schema_index)
                .unwrap_or(&CellValue::Empty)
                .clone(),
            skip: false,
        };
        before_csv_cell(handlers, &mut context)?;
        if !context.skip {
            record[*physical_index] = context.value.as_text();
        }
        after_csv_cell(handlers, &context)?;
    }
    after_csv_row(handlers, &row_context)?;
    Ok(record)
}

fn csv_record(columns: &[(usize, usize, &'static ExcelColumn)]) -> Vec<String> {
    vec![
        String::new();
        columns
            .iter()
            .map(|(physical_index, _, _)| physical_index + 1)
            .max()
            .unwrap_or(0)
    ]
}

fn before_csv_row(handlers: &mut [Box<dyn WriteHandler>], context: &WriteRowContext) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.before_row(context)?;
    }
    Ok(())
}

fn after_csv_row(handlers: &mut [Box<dyn WriteHandler>], context: &WriteRowContext) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.after_row(context)?;
    }
    Ok(())
}

fn before_csv_cell(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &mut WriteCellContext,
) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.before_cell(context)?;
    }
    Ok(())
}

fn after_csv_cell(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteCellContext,
) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.after_cell(context)?;
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct WriteProgress {
    next_row: u32,
    next_data_index: usize,
}

#[derive(Clone, Copy)]
struct SheetStyleContext<'a> {
    explicit: Option<&'a CellStyle>,
    metadata: &'a ExcelWriteMetadata,
    is_head: bool,
}

impl<'a> SheetStyleContext<'a> {
    const fn head(explicit: &'a CellStyle, metadata: &'a ExcelWriteMetadata) -> Self {
        Self {
            explicit: Some(explicit),
            metadata,
            is_head: true,
        }
    }

    const fn content(explicit: Option<&'a CellStyle>, metadata: &'a ExcelWriteMetadata) -> Self {
        Self {
            explicit,
            metadata,
            is_head: false,
        }
    }

    const fn column(self, column: &'a ExcelColumn) -> CellFormatContext<'a> {
        let (cell, font) = if self.is_head {
            (
                match column.head_style {
                    Some(style) => Some(style),
                    None => self.metadata.head_style,
                },
                match column.head_font_style {
                    Some(style) => Some(style),
                    None => self.metadata.head_font_style,
                },
            )
        } else {
            (
                match column.content_style {
                    Some(style) => Some(style),
                    None => self.metadata.content_style,
                },
                match column.content_font_style {
                    Some(style) => Some(style),
                    None => self.metadata.content_font_style,
                },
            )
        };
        CellFormatContext {
            explicit: self.explicit,
            cell,
            font,
        }
    }
}

#[derive(Clone, Copy)]
struct CellFormatContext<'a> {
    explicit: Option<&'a CellStyle>,
    cell: Option<ExcelCellStyle>,
    font: Option<ExcelFontStyle>,
}

#[derive(Debug)]
struct ImageLayout {
    column_widths: HashMap<u16, u32>,
    head_rows: u32,
    head_row_height: u32,
    content_row_height: u32,
}

impl Default for ImageLayout {
    fn default() -> Self {
        Self {
            column_widths: HashMap::new(),
            head_rows: 0,
            head_row_height: 20,
            content_row_height: 20,
        }
    }
}

impl ImageLayout {
    fn new(
        columns: &[(usize, usize, &'static ExcelColumn)],
        options: &WriteOptions,
        metadata: &ExcelWriteMetadata,
        head_rows: u32,
    ) -> Result<Self> {
        let mut column_widths = HashMap::new();
        for (column, width) in &options.column_widths {
            column_widths.insert(*column, excel_column_width_pixels(*width));
        }
        for (physical_index, _, column) in columns {
            let physical_index = to_column(*physical_index)?;
            if column_widths.contains_key(&physical_index) {
                continue;
            }
            if let Some(width) = column.column_width.or(metadata.column_width) {
                column_widths.insert(physical_index, excel_column_width_pixels(width));
            }
        }
        Ok(Self {
            column_widths,
            head_rows,
            head_row_height: excel_row_height_pixels(metadata.head_row_height),
            content_row_height: excel_row_height_pixels(metadata.content_row_height),
        })
    }

    fn column_width(&self, column: u16) -> u32 {
        self.column_widths.get(&column).copied().unwrap_or(64)
    }

    const fn row_height(&self, row: u32) -> u32 {
        if row < self.head_rows {
            self.head_row_height
        } else {
            self.content_row_height
        }
    }
}

fn excel_column_width_pixels(width: u16) -> u32 {
    if width == 0 {
        0
    } else {
        u32::from(width) * 7 + 5
    }
}

fn excel_row_height_pixels(height: Option<u16>) -> u32 {
    height.map_or(20, |height| (u32::from(height) * 4 + 1) / 3)
}

fn write_sheet_to_workbook<T, I>(
    workbook: &mut Workbook,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let worksheet = if options.constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(&options.sheet_name)
        .map_err(format_error)?;
    for (column, width) in &options.column_widths {
        worksheet
            .set_column_width(*column, *width)
            .map_err(format_error)?;
    }
    apply_annotation_column_widths::<T>(worksheet, options)?;
    for range in &options.merge_ranges {
        worksheet
            .merge_range(
                range.first_row,
                range.first_column,
                range.last_row,
                range.last_column,
                "",
                &Format::new(),
            )
            .map_err(format_error)?;
    }
    let head_rows = head_rows_for_schema(T::schema(), options)?;
    let freeze_panes = options
        .freeze_panes
        .or_else(|| (options.freeze_head && options.need_head).then_some((head_rows, 0)));
    if let Some((row, column)) = freeze_panes {
        worksheet
            .set_freeze_panes(row, column)
            .map_err(format_error)?;
    }

    let sheet_context = WriteSheetContext::new(&options.sheet_name);
    before_sheet(handlers, &sheet_context)?;

    let progress = append_rows_to_worksheet::<T, I>(
        worksheet,
        options,
        rows,
        handlers,
        WriteProgress {
            next_row: 0,
            next_data_index: 0,
        },
        true,
        T::write_metadata(),
    )?;
    after_sheet(handlers, &sheet_context)?;
    if options.auto_width {
        worksheet.autofit();
    }
    Ok(progress)
}

fn append_rows_to_worksheet<T, I>(
    worksheet: &mut Worksheet,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
    progress: WriteProgress,
    write_head: bool,
    metadata: &ExcelWriteMetadata,
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let WriteProgress {
        next_row: mut row_index,
        next_data_index: mut data_index,
    } = progress;
    let columns = selected_columns(T::schema(), options);
    let head_rows = head_rows_for_schema(T::schema(), options)?;
    let image_layout = ImageLayout::new(&columns, options, metadata, head_rows)?;
    if write_head && head_rows > 0 {
        if let Some(head) = &options.dynamic_head {
            write_dynamic_headers_with_handlers(
                worksheet,
                &columns,
                head,
                &options.sheet_name,
                SheetStyleContext::head(&options.head_style, metadata),
                handlers,
                &image_layout,
            )?;
        } else {
            write_headers_with_handlers(
                worksheet,
                &columns,
                &options.sheet_name,
                SheetStyleContext::head(&options.head_style, metadata),
                handlers,
                &image_layout,
            )?;
        }
        if let Some(height) = metadata.head_row_height {
            for head_row in row_index..row_index + head_rows {
                worksheet
                    .set_row_height(head_row, height)
                    .map_err(format_error)?;
            }
        }
        row_index += head_rows;
    }
    for row in rows {
        if let Some(height) = metadata.content_row_height {
            worksheet
                .set_row_height(row_index, height)
                .map_err(format_error)?;
        }
        let cells = row.to_row_with_converters(&options.converters)?;
        let dynamic_columns = dynamic_columns_for_row(T::schema().is_empty(), cells.len(), options);
        let row_columns = dynamic_columns.as_deref().unwrap_or(&columns);
        let style = (!options.content_styles.is_empty())
            .then(|| &options.content_styles[data_index % options.content_styles.len()]);
        apply_loop_merges(worksheet, row_index, data_index, &options.loop_merges)?;
        write_data_row_with_handlers(
            worksheet,
            row_index,
            row_columns,
            &cells,
            &options.sheet_name,
            SheetStyleContext::content(style, metadata),
            handlers,
            &image_layout,
        )?;
        row_index += 1;
        data_index += 1;
    }
    Ok(WriteProgress {
        next_row: row_index,
        next_data_index: data_index,
    })
}

fn apply_loop_merges(
    worksheet: &mut Worksheet,
    row_index: u32,
    data_index: usize,
    strategies: &[LoopMergeStrategy],
) -> Result<()> {
    for strategy in strategies {
        #[allow(clippy::cast_possible_truncation)]
        let each_rows = strategy.each_rows as usize;
        if !data_index.is_multiple_of(each_rows) {
            continue;
        }
        let last_row = row_index
            .checked_add(strategy.each_rows - 1)
            .ok_or_else(|| ExcelError::Format("loop merge row overflow".to_owned()))?;
        let last_column = strategy
            .column_index
            .checked_add(strategy.column_extend - 1)
            .ok_or_else(|| ExcelError::Format("loop merge column overflow".to_owned()))?;
        worksheet
            .merge_range(
                row_index,
                strategy.column_index,
                last_row,
                last_column,
                "",
                &Format::new(),
            )
            .map_err(format_error)?;
    }
    Ok(())
}

fn sort_handlers(handlers: &mut [Box<dyn WriteHandler>]) {
    handlers.sort_by_key(|handler| handler.order());
}

fn before_workbook(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.before_workbook(context)?;
    }
    Ok(())
}

fn after_workbook(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.after_workbook(context)?;
    }
    Ok(())
}

fn before_sheet(handlers: &mut [Box<dyn WriteHandler>], context: &WriteSheetContext) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.before_sheet(context)?;
    }
    Ok(())
}

fn after_sheet(handlers: &mut [Box<dyn WriteHandler>], context: &WriteSheetContext) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.after_sheet(context)?;
    }
    Ok(())
}

fn ordered_columns(schema: &'static [ExcelColumn]) -> Vec<(usize, usize, &'static ExcelColumn)> {
    let mut columns = schema
        .iter()
        .enumerate()
        .map(|(schema_index, column)| {
            let physical_index = column.index.unwrap_or(schema_index);
            (physical_index, schema_index, column)
        })
        .collect::<Vec<_>>();
    columns.sort_by_key(|(physical_index, schema_index, column)| {
        (*physical_index, column.order, *schema_index)
    });
    columns
}

fn apply_annotation_column_widths<T>(
    worksheet: &mut Worksheet,
    options: &WriteOptions,
) -> Result<()>
where
    T: ExcelRow,
{
    let type_width = T::write_metadata().column_width;
    for (physical_index, _, column) in selected_columns(T::schema(), options) {
        if options
            .column_widths
            .iter()
            .any(|(explicit, _)| usize::from(*explicit) == physical_index)
        {
            continue;
        }
        if let Some(width) = column.column_width.or(type_width) {
            worksheet
                .set_column_width(to_column(physical_index)?, width)
                .map_err(format_error)?;
        }
    }
    Ok(())
}

fn selected_columns(
    schema: &'static [ExcelColumn],
    options: &WriteOptions,
) -> Vec<(usize, usize, &'static ExcelColumn)> {
    if schema.is_empty()
        && let Some(head) = &options.dynamic_head
    {
        return selected_dynamic_columns(head.len(), options);
    }
    let mut columns = ordered_columns(schema)
        .into_iter()
        .filter(|(physical_index, _, column)| {
            let included_by_index = options
                .include_column_indexes
                .as_ref()
                .is_some_and(|indexes| indexes.contains(physical_index));
            let included_by_name = options
                .include_column_field_names
                .as_ref()
                .is_some_and(|names| names.iter().any(|name| name == column.field));
            let has_includes = options.include_column_indexes.is_some()
                || options.include_column_field_names.is_some();
            let excluded = options.exclude_column_indexes.contains(physical_index)
                || options
                    .exclude_column_field_names
                    .iter()
                    .any(|name| name == column.field);
            (!has_includes || included_by_index || included_by_name) && !excluded
        })
        .collect::<Vec<_>>();

    if options.order_by_include_column {
        columns.sort_by_key(|(physical_index, _, column)| {
            options
                .include_column_indexes
                .as_ref()
                .and_then(|indexes| indexes.iter().position(|index| index == physical_index))
                .or_else(|| {
                    options
                        .include_column_field_names
                        .as_ref()
                        .and_then(|names| names.iter().position(|name| name == column.field))
                })
                .unwrap_or(usize::MAX)
        });
        for (output_index, (physical_index, _, _)) in columns.iter_mut().enumerate() {
            *physical_index = output_index;
        }
    }
    columns
}

const DYNAMIC_COLUMN: ExcelColumn = ExcelColumn::new("", "", None, i32::MAX, None);

#[inline(never)]
fn selected_dynamic_columns(
    column_count: usize,
    options: &WriteOptions,
) -> Vec<(usize, usize, &'static ExcelColumn)> {
    let mut columns = Vec::with_capacity(column_count);
    for index in 0..column_count {
        let included_by_index = match &options.include_column_indexes {
            Some(indexes) => indexes.contains(&index),
            None => false,
        };
        let has_includes = options.include_column_indexes.is_some()
            || options.include_column_field_names.is_some();
        let excluded = options.exclude_column_indexes.contains(&index);
        if (!has_includes || included_by_index) && !excluded {
            columns.push((index, index, &DYNAMIC_COLUMN));
        }
    }

    if options.order_by_include_column {
        if let Some(indexes) = &options.include_column_indexes {
            let mut ordered = Vec::with_capacity(columns.len());
            for requested in indexes {
                for column in &columns {
                    if column.1 == *requested {
                        ordered.push(*column);
                        break;
                    }
                }
            }
            columns = ordered;
        }
        for (output_index, (physical_index, _, _)) in columns.iter_mut().enumerate() {
            *physical_index = output_index;
        }
    }
    columns
}

fn dynamic_columns_for_row(
    schema_is_empty: bool,
    column_count: usize,
    options: &WriteOptions,
) -> Option<Vec<(usize, usize, &'static ExcelColumn)>> {
    (schema_is_empty && options.dynamic_head.is_none())
        .then(|| selected_dynamic_columns(column_count, options))
}

fn head_rows_for_schema(schema: &[ExcelColumn], options: &WriteOptions) -> Result<u32> {
    head_rows_for_schema_state(schema.is_empty(), options)
}

fn head_rows_for_schema_state(schema_is_empty: bool, options: &WriteOptions) -> Result<u32> {
    if schema_is_empty && options.dynamic_head.is_none() {
        return Ok(0);
    }
    dynamic_head_rows(options)
}

fn dynamic_head_rows(options: &WriteOptions) -> Result<u32> {
    if !options.need_head {
        return Ok(0);
    }
    let Some(head) = &options.dynamic_head else {
        return Ok(1);
    };
    if head.is_empty() || head.iter().any(Vec::is_empty) {
        return Err(ExcelError::Format(
            "dynamic head must contain at least one non-empty path".to_owned(),
        ));
    }
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    head_level_to_row(levels)
}

fn head_level_to_row(level: usize) -> Result<u32> {
    u32::try_from(level).map_err(|_| ExcelError::Format("dynamic head is too deep".to_owned()))
}

#[cfg(test)]
fn write_headers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
) -> Result<()> {
    const METADATA: ExcelWriteMetadata = ExcelWriteMetadata::new();
    let layout = ImageLayout::default();
    write_headers_with_handlers(
        worksheet,
        columns,
        "",
        SheetStyleContext::head(&CellStyle::new(), &METADATA),
        &mut [],
        &layout,
    )
}

fn write_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    sheet_name: &str,
    style: SheetStyleContext<'_>,
    handlers: &mut [Box<dyn WriteHandler>],
    image_layout: &ImageLayout,
) -> Result<()> {
    let labels = columns
        .iter()
        .map(|(_, _, column)| column.name.to_owned())
        .collect::<Vec<_>>();
    write_header_row_with_handlers(
        worksheet,
        0,
        columns,
        &labels,
        sheet_name,
        style,
        handlers,
        image_layout,
    )
}

fn write_dynamic_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    head: &[Vec<String>],
    sheet_name: &str,
    style: SheetStyleContext<'_>,
    handlers: &mut [Box<dyn WriteHandler>],
    image_layout: &ImageLayout,
) -> Result<()> {
    if head.len() != columns.len() {
        return Err(ExcelError::Format(format!(
            "dynamic head column count {} does not match selected column count {}",
            head.len(),
            columns.len()
        )));
    }
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    for level in 0..levels {
        #[allow(clippy::cast_possible_truncation)]
        let row_index = level as u32;
        let labels = head
            .iter()
            .map(|path| path.get(level).cloned().unwrap_or_default())
            .collect::<Vec<_>>();
        write_header_row_with_handlers(
            worksheet,
            row_index,
            columns,
            &labels,
            sheet_name,
            style,
            handlers,
            image_layout,
        )?;
    }
    merge_dynamic_head_groups(worksheet, columns, head, style)
}

#[allow(clippy::too_many_arguments)]
fn write_header_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    labels: &[String],
    sheet_name: &str,
    style: SheetStyleContext<'_>,
    handlers: &mut [Box<dyn WriteHandler>],
    image_layout: &ImageLayout,
) -> Result<()> {
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: true,
    };
    for handler in handlers.iter_mut() {
        handler.before_row(&row_context)?;
    }
    for ((physical_index, _, column), label) in columns.iter().zip(labels) {
        let format_context = style.column(column);
        let format = cell_format(format_context);
        let column_index = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index,
            field: (!column.field.is_empty()).then_some(column.field),
            is_head: true,
            value: CellValue::String(label.clone()),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            match &context.value {
                CellValue::String(value) | CellValue::Error(value) => {
                    worksheet
                        .write_string_with_format(row_index, context.column_index, value, &format)
                        .map_err(format_error)?;
                }
                value => write_cell(
                    worksheet,
                    row_index,
                    context.column_index,
                    column,
                    value,
                    format_context,
                    image_layout,
                )?,
            }
        }
        for handler in handlers.iter_mut() {
            handler.after_cell(&context)?;
        }
    }
    for handler in handlers.iter_mut() {
        handler.after_row(&row_context)?;
    }
    Ok(())
}

fn merge_dynamic_head_groups(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    head: &[Vec<String>],
    style: SheetStyleContext<'_>,
) -> Result<()> {
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    for level in 0..levels {
        #[allow(clippy::cast_possible_truncation)]
        let row_index = level as u32;
        let mut start = 0;
        while start < head.len() {
            let mut end = start;
            while end + 1 < head.len()
                && columns[end].0.checked_add(1) == Some(columns[end + 1].0)
                && same_dynamic_head_group(head, start, end + 1, level)
            {
                end += 1;
            }
            let label = head[start].get(level).map_or("", String::as_str);
            if end > start && !label.is_empty() {
                let format = cell_format(style.column(columns[start].2));
                worksheet
                    .merge_range(
                        row_index,
                        to_column(columns[start].0)?,
                        row_index,
                        to_column(columns[end].0)?,
                        label,
                        &format,
                    )
                    .map_err(format_error)?;
            }
            start = end + 1;
        }
    }
    Ok(())
}

fn same_dynamic_head_group(
    head: &[Vec<String>],
    first: usize,
    second: usize,
    level: usize,
) -> bool {
    head[first].get(level) == head[second].get(level)
        && head[first]
            .iter()
            .take(level)
            .eq(head[second].iter().take(level))
}

#[cfg(test)]
fn write_data_row(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
) -> Result<()> {
    let image_layout = ImageLayout::default();
    write_data_row_with_handlers(
        worksheet,
        row_index,
        columns,
        cells,
        "",
        SheetStyleContext {
            explicit: None,
            metadata: &ExcelWriteMetadata::new(),
            is_head: false,
        },
        &mut [],
        &image_layout,
    )
}

#[allow(clippy::too_many_arguments)]
fn write_data_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
    sheet_name: &str,
    style: SheetStyleContext<'_>,
    handlers: &mut [Box<dyn WriteHandler>],
    image_layout: &ImageLayout,
) -> Result<()> {
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: false,
    };
    for handler in handlers.iter_mut() {
        handler.before_row(&row_context)?;
    }
    for (physical_index, schema_index, metadata) in columns {
        let value = cells.get(*schema_index).unwrap_or(&CellValue::Empty);
        let column = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index: column,
            field: (!metadata.field.is_empty()).then_some(metadata.field),
            is_head: false,
            value: value.clone(),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            let format_context = style.column(metadata);
            write_cell(
                worksheet,
                row_index,
                context.column_index,
                metadata,
                &context.value,
                format_context,
                image_layout,
            )?;
        }
        for handler in handlers.iter_mut() {
            handler.after_cell(&context)?;
        }
    }
    for handler in handlers.iter_mut() {
        handler.after_row(&row_context)?;
    }
    Ok(())
}

fn write_cell(
    worksheet: &mut Worksheet,
    row_index: u32,
    column: u16,
    metadata: &ExcelColumn,
    value: &CellValue,
    style: CellFormatContext<'_>,
    image_layout: &ImageLayout,
) -> Result<()> {
    let format = cell_format(style);
    match value {
        CellValue::Empty => {
            worksheet
                .write_blank(row_index, column, &format)
                .map_err(format_error)?;
        }
        CellValue::String(value) | CellValue::Error(value) => {
            worksheet
                .write_string_with_format(row_index, column, value, &format)
                .map_err(format_error)?;
        }
        CellValue::Bool(value) => {
            worksheet
                .write_boolean_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::Int(value) => {
            write_integer(worksheet, row_index, column, *value, &format)?;
        }
        CellValue::Float(value) => {
            worksheet
                .write_number_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::Decimal(value) => {
            let value = value
                .to_f64()
                .filter(|value| value.is_finite())
                .ok_or_else(|| {
                    ExcelError::Format("decimal value exceeds XLSX numeric range".to_owned())
                })?;
            worksheet
                .write_number_with_format(row_index, column, value, &format)
                .map_err(format_error)?;
        }
        CellValue::Date(value) => {
            let format = format
                .clone()
                .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd"));
            worksheet
                .write_datetime_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::DateTime(value) => {
            let format = format
                .clone()
                .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd hh:mm:ss"));
            worksheet
                .write_datetime_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::Formula(value) => {
            worksheet
                .write_formula_with_format(row_index, column, value.as_str(), &format)
                .map_err(format_error)?;
        }
        CellValue::Hyperlink { url, text } => {
            worksheet
                .write_url_with_options(row_index, column, url.as_str(), text, "", Some(&format))
                .map_err(format_error)?;
        }
        CellValue::Comment { value, text } => {
            write_cell(
                worksheet,
                row_index,
                column,
                metadata,
                value,
                style,
                image_layout,
            )?;
            worksheet
                .insert_note(row_index, column, &Note::new(text))
                .map_err(format_error)?;
        }
        CellValue::Image(bytes) => {
            let image = image_from_buffer(bytes)?;
            worksheet
                .insert_image_fit_to_cell(row_index, column, &image, true)
                .map_err(format_error)?;
        }
        CellValue::Images { value, images } => {
            write_cell(
                worksheet,
                row_index,
                column,
                metadata,
                value,
                style,
                image_layout,
            )?;
            for image in images {
                insert_image_data(worksheet, row_index, column, image, image_layout)?;
            }
        }
    }
    Ok(())
}

fn image_from_buffer(bytes: &[u8]) -> Result<Image> {
    if bytes.len() < 8 {
        return Err(ExcelError::Format(
            "image buffer is too short to contain a valid header".to_owned(),
        ));
    }
    Image::new_from_buffer(bytes).map_err(format_error)
}

fn insert_image_data(
    worksheet: &mut Worksheet,
    current_row: u32,
    current_column: u16,
    data: &ImageData,
    layout: &ImageLayout,
) -> Result<()> {
    let anchor = data.get_anchor();
    let coordinates = anchor.get_coordinates();
    let first_row = resolve_anchor_coordinate(
        current_row,
        coordinates.get_first_row_index(),
        coordinates.get_relative_first_row_index(),
        "first row",
    )?;
    let first_column = resolve_anchor_coordinate(
        u32::from(current_column),
        coordinates.get_first_column_index().map(u32::from),
        coordinates.get_relative_first_column_index(),
        "first column",
    )?;
    let last_row = resolve_anchor_coordinate(
        current_row,
        coordinates.get_last_row_index(),
        coordinates.get_relative_last_row_index(),
        "last row",
    )?;
    let last_column = resolve_anchor_coordinate(
        u32::from(current_column),
        coordinates.get_last_column_index().map(u32::from),
        coordinates.get_relative_last_column_index(),
        "last column",
    )?;
    if first_row > last_row || first_column > last_column {
        return Err(ExcelError::Format(
            "image anchor start must not follow its end".to_owned(),
        ));
    }
    let first_column = u16::try_from(first_column)
        .map_err(|_| ExcelError::Format("image anchor column exceeds XLSX limit".to_owned()))?;
    let last_column = u16::try_from(last_column)
        .map_err(|_| ExcelError::Format("image anchor column exceeds XLSX limit".to_owned()))?;
    if last_row >= 1_048_576 || last_column >= 16_384 {
        return Err(ExcelError::Format(
            "image anchor exceeds XLSX worksheet limits".to_owned(),
        ));
    }

    let total_width = (first_column..=last_column).try_fold(0_u32, |width, column| {
        width
            .checked_add(layout.column_width(column))
            .ok_or_else(|| ExcelError::Format("image anchor width overflow".to_owned()))
    })?;
    let total_height = (first_row..=last_row).try_fold(0_u32, |height, row| {
        height
            .checked_add(layout.row_height(row))
            .ok_or_else(|| ExcelError::Format("image anchor height overflow".to_owned()))
    })?;
    let left = anchor.get_left().unwrap_or(0);
    let right = anchor.get_right().unwrap_or(0);
    let top = anchor.get_top().unwrap_or(0);
    let bottom = anchor.get_bottom().unwrap_or(0);
    let width = total_width
        .checked_sub(left)
        .and_then(|value| value.checked_sub(right))
        .filter(|value| *value > 0)
        .ok_or_else(|| {
            ExcelError::Format("image horizontal margins consume its anchor".to_owned())
        })?;
    let height = total_height
        .checked_sub(top)
        .and_then(|value| value.checked_sub(bottom))
        .filter(|value| *value > 0)
        .ok_or_else(|| {
            ExcelError::Format("image vertical margins consume its anchor".to_owned())
        })?;
    let movement = match anchor
        .get_anchor_type()
        .unwrap_or(AnchorType::MoveAndResize)
    {
        AnchorType::MoveAndResize => ObjectMovement::MoveAndSizeWithCells,
        AnchorType::DontMoveDoResize | AnchorType::MoveDontResize => {
            ObjectMovement::MoveButDontSizeWithCells
        }
        AnchorType::DontMoveAndResize => ObjectMovement::DontMoveOrSizeWithCells,
    };
    let image = image_from_buffer(data.image())?
        .set_scale_to_size(width, height, false)
        .set_object_movement(movement);
    insert_scaled_image(worksheet, first_row, first_column, &image, left, top)
}

fn insert_scaled_image(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    image: &Image,
    left: u32,
    top: u32,
) -> Result<()> {
    worksheet
        .insert_image_with_offset(row, column, image, left, top)
        .map(|_| ())
        .map_err(format_error)
}

fn resolve_anchor_coordinate(
    current: u32,
    absolute: Option<u32>,
    relative: Option<i32>,
    label: &str,
) -> Result<u32> {
    if let Some(absolute) = absolute.filter(|value| *value > 0) {
        return Ok(absolute);
    }
    let Some(relative) = relative else {
        return Ok(current);
    };
    current
        .checked_add_signed(relative)
        .ok_or_else(|| ExcelError::Format(format!("image anchor {label} is outside the worksheet")))
}

fn cell_format(context: CellFormatContext<'_>) -> Format {
    let mut format = Format::new();
    if let Some(style) = context.cell {
        format = apply_annotation_cell_style(format, style);
    }
    if let Some(font) = context.font {
        format = apply_annotation_font_style(format, font);
    }
    let Some(style) = context.explicit else {
        return format;
    };
    if style.bold {
        format = format.set_bold();
    }
    if style.italic {
        format = format.set_italic();
    }
    if let Some(color) = style.font_color {
        format = format.set_font_color(color);
    }
    if let Some(color) = style.background_color {
        format = format
            .set_background_color(color)
            .set_pattern(FormatPattern::Solid);
    }
    if let Some(alignment) = style.horizontal_alignment {
        format = format.set_align(horizontal_format_align(alignment));
    }
    if let Some(alignment) = style.vertical_alignment {
        format = format.set_align(vertical_format_align(alignment));
    }
    if style.wrap_text {
        format = format.set_text_wrap();
    }
    if let Some(number_format) = &style.number_format {
        format = format.set_num_format(number_format);
    }
    format
}

fn apply_annotation_cell_style(mut format: Format, style: ExcelCellStyle) -> Format {
    if let Some(hidden) = style.hidden {
        format = if hidden {
            format.set_hidden()
        } else {
            format.unset_hidden()
        };
    }
    if let Some(locked) = style.locked {
        format = if locked {
            format.set_locked()
        } else {
            format.set_unlocked()
        };
    }
    if let Some(quote_prefix) = style.quote_prefix {
        format = if quote_prefix {
            format.set_quote_prefix()
        } else {
            format.unset_quote_prefix()
        };
    }
    if let Some(alignment) = style.horizontal_alignment {
        format = format.set_align(annotation_horizontal_format_align(alignment));
    }
    if let Some(wrapped) = style.wrapped {
        format = if wrapped {
            format.set_text_wrap()
        } else {
            format.unset_text_wrap()
        };
    }
    if let Some(alignment) = style.vertical_alignment {
        format = format.set_align(annotation_vertical_format_align(alignment));
    }
    if let Some(rotation) = style.rotation {
        format = format.set_rotation(rotation);
    }
    if let Some(indent) = style.indent {
        format = format.set_indent(indent);
    }
    if let Some(border) = style.border_left {
        format = format.set_border_left(annotation_border_style(border));
    }
    if let Some(border) = style.border_right {
        format = format.set_border_right(annotation_border_style(border));
    }
    if let Some(border) = style.border_top {
        format = format.set_border_top(annotation_border_style(border));
    }
    if let Some(border) = style.border_bottom {
        format = format.set_border_bottom(annotation_border_style(border));
    }
    if let Some(color) = style.left_border_color {
        format = format.set_border_left_color(annotation_color(color));
    }
    if let Some(color) = style.right_border_color {
        format = format.set_border_right_color(annotation_color(color));
    }
    if let Some(color) = style.top_border_color {
        format = format.set_border_top_color(annotation_color(color));
    }
    if let Some(color) = style.bottom_border_color {
        format = format.set_border_bottom_color(annotation_color(color));
    }
    if let Some(pattern) = style.fill_pattern {
        format = format.set_pattern(annotation_fill_pattern(pattern));
    }
    if let Some(color) = style.fill_background_color {
        format = format.set_background_color(annotation_color(color));
    }
    if let Some(color) = style.fill_foreground_color {
        format = format.set_foreground_color(annotation_color(color));
    }
    if let Some(shrink) = style.shrink_to_fit {
        format = if shrink {
            format.set_shrink()
        } else {
            format.unset_shrink()
        };
    }
    if let Some(data_format) = style.data_format {
        format = match data_format {
            ExcelDataFormat::Builtin(index) => format.set_num_format_index(index),
            ExcelDataFormat::Custom(value) => format.set_num_format(value),
        };
    }
    format
}

fn apply_annotation_font_style(mut format: Format, style: ExcelFontStyle) -> Format {
    if let Some(font_name) = style.font_name {
        format = format.set_font_name(font_name);
    }
    if let Some(font_height) = style.font_height_in_points {
        format = format.set_font_size(font_height);
    }
    if let Some(italic) = style.italic {
        format = if italic {
            format.set_italic()
        } else {
            format.unset_italic()
        };
    }
    if let Some(strikeout) = style.strikeout {
        format = if strikeout {
            format.set_font_strikethrough()
        } else {
            format.unset_font_strikethrough()
        };
    }
    if let Some(color) = style.color {
        format = format.set_font_color(annotation_color(color));
    }
    if let Some(script) = style.type_offset {
        format = format.set_font_script(annotation_font_script(script));
    }
    if let Some(underline) = style.underline {
        format = format.set_underline(annotation_underline(underline));
    }
    if let Some(charset) = style.charset {
        format = format.set_font_charset(charset);
    }
    if let Some(bold) = style.bold {
        format = if bold {
            format.set_bold()
        } else {
            format.unset_bold()
        };
    }
    format
}

fn annotation_color(color: ExcelColor) -> Color {
    match color {
        ExcelColor::Rgb(value) => Color::RGB(value),
        ExcelColor::Indexed(64) => Color::Automatic,
        ExcelColor::Indexed(index) => indexed_color(index),
    }
}

fn indexed_color(index: u8) -> Color {
    let rgb = match index {
        0 | 8 => 0x0000_0000,
        1 | 9 => 0x00ff_ffff,
        2 | 10 => 0x00ff_0000,
        3 | 11 => 0x0000_ff00,
        4 | 12 | 39 => 0x0000_00ff,
        5 | 13 | 34 => 0x00ff_ff00,
        6 | 14 | 33 => 0x00ff_00ff,
        7 | 15 | 35 => 0x0000_ffff,
        16 | 37 => 0x0080_0000,
        17 => 0x0000_8000,
        18 | 32 => 0x0000_0080,
        19 => 0x0080_8000,
        20 | 36 => 0x0080_0080,
        21 | 38 => 0x0000_8080,
        22 => 0x00c0_c0c0,
        23 => 0x0080_8080,
        24 => 0x0099_99ff,
        25 => 0x007f_0000,
        26 => 0x00ff_ffcc,
        27 | 41 => 0x00cc_ffff,
        28 => 0x0066_0066,
        29 => 0x00ff_8080,
        30 => 0x0000_66cc,
        31 => 0x00cc_ccff,
        40 => 0x0000_ccff,
        42 => 0x00cc_ffcc,
        43 => 0x00ff_ff99,
        44 => 0x0099_ccff,
        45 => 0x00ff_99cc,
        46 => 0x00cc_99ff,
        47 => 0x00ff_cc99,
        48 => 0x0033_66ff,
        49 => 0x0033_cccc,
        50 => 0x0099_cc00,
        51 => 0x00ff_cc00,
        52 => 0x00ff_9900,
        53 => 0x00ff_6600,
        54 => 0x0066_6699,
        55 => 0x0096_9696,
        56 => 0x0000_3366,
        57 => 0x0033_9966,
        58 => 0x0000_3300,
        59 => 0x0033_3300,
        60 => 0x0099_3300,
        61 => 0x0099_3366,
        62 => 0x0033_3399,
        63 => 0x0033_3333,
        _ => return Color::Default,
    };
    Color::RGB(rgb)
}

const fn annotation_horizontal_format_align(alignment: ExcelHorizontalAlignment) -> FormatAlign {
    match alignment {
        ExcelHorizontalAlignment::General => FormatAlign::General,
        ExcelHorizontalAlignment::Left => FormatAlign::Left,
        ExcelHorizontalAlignment::Center => FormatAlign::Center,
        ExcelHorizontalAlignment::Right => FormatAlign::Right,
        ExcelHorizontalAlignment::Fill => FormatAlign::Fill,
        ExcelHorizontalAlignment::Justify => FormatAlign::Justify,
        ExcelHorizontalAlignment::CenterAcross => FormatAlign::CenterAcross,
        ExcelHorizontalAlignment::Distributed => FormatAlign::Distributed,
    }
}

const fn annotation_vertical_format_align(alignment: ExcelVerticalAlignment) -> FormatAlign {
    match alignment {
        ExcelVerticalAlignment::Top => FormatAlign::Top,
        ExcelVerticalAlignment::Center => FormatAlign::VerticalCenter,
        ExcelVerticalAlignment::Bottom => FormatAlign::Bottom,
        ExcelVerticalAlignment::Justify => FormatAlign::VerticalJustify,
        ExcelVerticalAlignment::Distributed => FormatAlign::VerticalDistributed,
    }
}

const fn annotation_border_style(border: ExcelBorderStyle) -> FormatBorder {
    match border {
        ExcelBorderStyle::None => FormatBorder::None,
        ExcelBorderStyle::Thin => FormatBorder::Thin,
        ExcelBorderStyle::Medium => FormatBorder::Medium,
        ExcelBorderStyle::Dashed => FormatBorder::Dashed,
        ExcelBorderStyle::Dotted => FormatBorder::Dotted,
        ExcelBorderStyle::Thick => FormatBorder::Thick,
        ExcelBorderStyle::Double => FormatBorder::Double,
        ExcelBorderStyle::Hair => FormatBorder::Hair,
        ExcelBorderStyle::MediumDashed => FormatBorder::MediumDashed,
        ExcelBorderStyle::DashDot => FormatBorder::DashDot,
        ExcelBorderStyle::MediumDashDot => FormatBorder::MediumDashDot,
        ExcelBorderStyle::DashDotDot => FormatBorder::DashDotDot,
        ExcelBorderStyle::MediumDashDotDot => FormatBorder::MediumDashDotDot,
        ExcelBorderStyle::SlantDashDot => FormatBorder::SlantDashDot,
    }
}

const fn annotation_fill_pattern(pattern: ExcelFillPattern) -> FormatPattern {
    match pattern {
        ExcelFillPattern::None => FormatPattern::None,
        ExcelFillPattern::Solid => FormatPattern::Solid,
        ExcelFillPattern::MediumGray => FormatPattern::MediumGray,
        ExcelFillPattern::DarkGray => FormatPattern::DarkGray,
        ExcelFillPattern::LightGray => FormatPattern::LightGray,
        ExcelFillPattern::DarkHorizontal => FormatPattern::DarkHorizontal,
        ExcelFillPattern::DarkVertical => FormatPattern::DarkVertical,
        ExcelFillPattern::DarkDown => FormatPattern::DarkDown,
        ExcelFillPattern::DarkUp => FormatPattern::DarkUp,
        ExcelFillPattern::DarkGrid => FormatPattern::DarkGrid,
        ExcelFillPattern::DarkTrellis => FormatPattern::DarkTrellis,
        ExcelFillPattern::LightHorizontal => FormatPattern::LightHorizontal,
        ExcelFillPattern::LightVertical => FormatPattern::LightVertical,
        ExcelFillPattern::LightDown => FormatPattern::LightDown,
        ExcelFillPattern::LightUp => FormatPattern::LightUp,
        ExcelFillPattern::LightGrid => FormatPattern::LightGrid,
        ExcelFillPattern::LightTrellis => FormatPattern::LightTrellis,
        ExcelFillPattern::Gray125 => FormatPattern::Gray125,
        ExcelFillPattern::Gray0625 => FormatPattern::Gray0625,
    }
}

const fn annotation_underline(underline: ExcelUnderline) -> FormatUnderline {
    match underline {
        ExcelUnderline::None => FormatUnderline::None,
        ExcelUnderline::Single => FormatUnderline::Single,
        ExcelUnderline::Double => FormatUnderline::Double,
        ExcelUnderline::SingleAccounting => FormatUnderline::SingleAccounting,
        ExcelUnderline::DoubleAccounting => FormatUnderline::DoubleAccounting,
    }
}

const fn annotation_font_script(script: ExcelFontScript) -> FormatScript {
    match script {
        ExcelFontScript::None => FormatScript::None,
        ExcelFontScript::Superscript => FormatScript::Superscript,
        ExcelFontScript::Subscript => FormatScript::Subscript,
    }
}

const fn horizontal_format_align(alignment: HorizontalAlignment) -> FormatAlign {
    match alignment {
        HorizontalAlignment::General => FormatAlign::General,
        HorizontalAlignment::Left => FormatAlign::Left,
        HorizontalAlignment::Center => FormatAlign::Center,
        HorizontalAlignment::Right => FormatAlign::Right,
        HorizontalAlignment::Fill => FormatAlign::Fill,
        HorizontalAlignment::Justify => FormatAlign::Justify,
        HorizontalAlignment::CenterAcross => FormatAlign::CenterAcross,
    }
}

const fn vertical_format_align(alignment: VerticalAlignment) -> FormatAlign {
    match alignment {
        VerticalAlignment::Top => FormatAlign::Top,
        VerticalAlignment::Center => FormatAlign::VerticalCenter,
        VerticalAlignment::Bottom => FormatAlign::Bottom,
        VerticalAlignment::Justify => FormatAlign::VerticalJustify,
        VerticalAlignment::Distributed => FormatAlign::VerticalDistributed,
    }
}

fn write_integer(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: i64,
    format: &Format,
) -> Result<()> {
    const MAX_EXACT_EXCEL_INTEGER: u64 = 9_007_199_254_740_991;
    if value.unsigned_abs() <= MAX_EXACT_EXCEL_INTEGER {
        #[allow(clippy::cast_precision_loss)]
        let number = value as f64;
        worksheet
            .write_number_with_format(row, column, number, format)
            .map(|_| ())
            .map_err(format_error)
    } else {
        worksheet
            .write_string_with_format(row, column, value.to_string(), format)
            .map(|_| ())
            .map_err(format_error)
    }
}

fn excel_date_format(format: Option<&str>, default: &str) -> String {
    format
        .unwrap_or(default)
        .replace("%Y", "yyyy")
        .replace("%m", "mm")
        .replace("%d", "dd")
        .replace("%H", "hh")
        .replace("%M", "mm")
        .replace("%S", "ss")
}

fn to_column(index: usize) -> Result<u16> {
    u16::try_from(index)
        .map_err(|_| ExcelError::Format("column index exceeds XLSX limit".to_owned()))
}

fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[cfg(test)]
mod tests;
