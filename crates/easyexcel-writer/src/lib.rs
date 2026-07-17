//! XLSX writer backed by `rust_xlsxwriter`.

use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Read, Seek, Write};
use std::marker::PhantomData;
use std::path::{Path, PathBuf};

use easyexcel_core::{
    CellValue, CsvCharset, ExcelColumn, ExcelError, ExcelRow, ExcelWriteMetadata, Result,
    WriteCellContext, WriteHandler, WriteRowContext, WriteSheetContext, WriteWorkbookContext,
};
use encoding_rs::{CoderResult, Encoding, UTF_8, UTF_16BE, UTF_16LE};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rust_xlsxwriter::{Format, FormatAlign, FormatPattern, Image, Note, Workbook, Worksheet};

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

        let options = sheet.options().clone();
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
    let mut rows = rows.into_iter().map(|row| row.to_row());
    write_csv_records(path, output, options, &columns, &mut rows, handlers)
}

fn write_csv_records(
    path: &Path,
    output: Box<dyn Write>,
    options: &WriteOptions,
    columns: &[(usize, usize, &'static ExcelColumn)],
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
    append_csv_records(&mut writer, options, columns, rows, handlers, 0, 0, true)?;
    finish_csv_record_writer(writer)?;
    after_sheet(handlers, &sheet_context)?;
    after_workbook(handlers, &workbook_context)
}

#[allow(clippy::too_many_arguments)]
fn append_csv_records(
    writer: &mut csv::Writer<CsvEncodingWriter>,
    options: &WriteOptions,
    columns: &[(usize, usize, &'static ExcelColumn)],
    rows: &mut dyn Iterator<Item = Result<Vec<CellValue>>>,
    handlers: &mut [Box<dyn WriteHandler>],
    mut row_index: u32,
    mut data_index: usize,
    write_head: bool,
) -> Result<WriteProgress> {
    let head_rows = dynamic_head_rows(options)?;
    if write_head && options.need_head {
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
        let record = csv_data_record(row_index, columns, &cells?, &options.sheet_name, handlers)?;
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
    let mut rows = rows.into_iter().map(|row| row.to_row());
    append_csv_records(
        writer, options, &columns, &mut rows, handlers, row_index, data_index, write_head,
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
            field: Some(column.field),
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
            field: Some(metadata.field),
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
    let head_rows = dynamic_head_rows(options)?;
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
    if write_head && options.need_head {
        if let Some(head) = &options.dynamic_head {
            write_dynamic_headers_with_handlers(
                worksheet,
                &columns,
                head,
                &options.sheet_name,
                &options.head_style,
                handlers,
            )?;
        } else {
            write_headers_with_handlers(
                worksheet,
                &columns,
                &options.sheet_name,
                &options.head_style,
                handlers,
            )?;
        }
        let head_rows = dynamic_head_rows(options)?;
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
        let cells = row.to_row()?;
        let style = (!options.content_styles.is_empty())
            .then(|| &options.content_styles[data_index % options.content_styles.len()]);
        apply_loop_merges(worksheet, row_index, data_index, &options.loop_merges)?;
        write_data_row_with_handlers(
            worksheet,
            row_index,
            &columns,
            &cells,
            &options.sheet_name,
            style,
            handlers,
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
    write_headers_with_handlers(worksheet, columns, "", &CellStyle::default(), &mut [])
}

fn write_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    sheet_name: &str,
    style: &CellStyle,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    let labels = columns
        .iter()
        .map(|(_, _, column)| column.name.to_owned())
        .collect::<Vec<_>>();
    write_header_row_with_handlers(worksheet, 0, columns, &labels, sheet_name, style, handlers)
}

fn write_dynamic_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    head: &[Vec<String>],
    sheet_name: &str,
    style: &CellStyle,
    handlers: &mut [Box<dyn WriteHandler>],
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
            worksheet, row_index, columns, &labels, sheet_name, style, handlers,
        )?;
    }
    merge_dynamic_head_groups(worksheet, columns, head, style)
}

fn write_header_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    labels: &[String],
    sheet_name: &str,
    style: &CellStyle,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    let format = cell_format(Some(style));
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: true,
    };
    for handler in handlers.iter_mut() {
        handler.before_row(&row_context)?;
    }
    for ((physical_index, _, column), label) in columns.iter().zip(labels) {
        let column_index = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index,
            field: Some(column.field),
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
                    Some(style),
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
    style: &CellStyle,
) -> Result<()> {
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    let format = cell_format(Some(style));
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
    write_data_row_with_handlers(worksheet, row_index, columns, cells, "", None, &mut [])
}

fn write_data_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
    sheet_name: &str,
    style: Option<&CellStyle>,
    handlers: &mut [Box<dyn WriteHandler>],
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
            field: Some(metadata.field),
            is_head: false,
            value: value.clone(),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            write_cell(
                worksheet,
                row_index,
                context.column_index,
                metadata,
                &context.value,
                style,
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
    style: Option<&CellStyle>,
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
            write_cell(worksheet, row_index, column, metadata, value, style)?;
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

fn cell_format(style: Option<&CellStyle>) -> Format {
    let Some(style) = style else {
        return Format::new();
    };
    let mut format = Format::new();
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
