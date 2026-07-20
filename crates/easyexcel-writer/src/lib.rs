//! XLSX writer backed by `rust_xlsxwriter`.

use std::collections::HashMap;
use std::fs::{File, OpenOptions};
use std::io::{Read, Seek, Write};
use std::marker::PhantomData;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use bigdecimal::ToPrimitive;
use easyexcel_core::{
    AnchorType, CacheLocation, CellValue, Converter, ConverterRegistry, CsvCharset, ExcelBorderStyle,
    ExcelCellStyle, ExcelColor, ExcelColumn, ExcelDataFormat, ExcelError, ExcelFillPattern,
    ExcelFontScript, ExcelFontStyle, ExcelHorizontalAlignment, ExcelRow, ExcelUnderline,
    ExcelVerticalAlignment, ExcelWriteMetadata, ImageData, Result, RichTextStringData,
    WriteCellContext, WriteFont, WriteHandler, WriteRowContext, WriteSheetContext,
    WriteWorkbookContext,
};
use encoding_rs::{CoderResult, Encoding, UTF_8, UTF_16BE, UTF_16LE};
use ms_offcrypto_writer::Ecma376AgileWriter;
use rust_xlsxwriter::{
    Color, Format, FormatAlign, FormatBorder, FormatPattern, FormatScript, FormatUnderline, Image,
    Note, ObjectMovement, Workbook, Worksheet,
};

// ---------------------------------------------------------------------------
// Mirrored Java sub-packages
// ---------------------------------------------------------------------------
pub mod builder;
mod excel_builder;
pub mod global_configuration;
pub mod executor;
pub mod handler;
pub mod holder;
pub mod merge;
pub mod metadata;
pub mod property;
pub mod style;
mod template_write;
/// SXSSF `GZIPSheetDataWriter` equivalent — gzip row spill for `compress_temp_files`.
pub mod gzip_spill;
/// BIFF8 (`.xls`) writer — Java `ExcelTypeEnum.XLS` / POI HSSF subset.
pub mod biff8;

pub use builder::abstract_excel_writer_parameter_builder::AbstractExcelWriterParameterBuilder;
pub use builder::excel_writer_table_builder::ExcelWriterTableBuilder;
pub use global_configuration::{
    apply_global_configuration_to_write_options, global_configuration_from_write_options,
};
pub use excel_builder::{ExcelBuilder, ExcelBuilderImpl, FillConfig as BuilderFillConfig};
pub use gzip_spill::{file_has_gzip_magic, GzipSpillSnapshot, GZIP_MAGIC};
pub use executor::abstract_excel_write_executor::AbstractExcelWriteExecutor;
pub use executor::excel_write_add_executor::ExcelWriteAddExecutor;
pub use executor::excel_write_executor::ExcelWriteExecutor;
pub use executor::excel_write_fill_executor::ExcelWriteFillExecutor;
pub use handler::abstract_cell_write_handler::AbstractCellWriteHandler;
pub use handler::abstract_row_write_handler::AbstractRowWriteHandler;
pub use handler::abstract_sheet_write_handler::AbstractSheetWriteHandler;
pub use handler::abstract_workbook_write_handler::AbstractWorkbookWriteHandler;
pub use handler::cell_write_handler::CellWriteHandler;
pub use handler::default_write_handler_loader::DefaultWriteHandlerLoader;
pub use handler::r#impl::impl_default_row_write_handler::{
    new_default_row_write_handler, DefaultRowWriteHandler,
};
pub use handler::r#impl::impl_dimension_workbook_write_handler::DimensionWorkbookWriteHandler;
pub use handler::r#impl::impl_fill_style_cell_write_handler::FillStyleCellWriteHandler;
pub use handler::row_write_handler::RowWriteHandler;
pub use handler::sheet_write_handler::SheetWriteHandler;
pub use handler::workbook_write_handler::WorkbookWriteHandler;
pub use holder::abstract_write_holder::AbstractWriteHolder;
pub use holder::write_holder::WriteHolder;
pub use holder::write_sheet_holder::WriteSheetHolder as MirroredWriteSheetHolder;
pub use holder::write_table_holder::WriteTableHolder as MirroredWriteTableHolder;
pub use holder::write_workbook_holder::WriteWorkbookHolder as MirroredWriteWorkbookHolder;
pub use merge::abstract_merge_strategy::AbstractMergeStrategy;
pub use merge::loop_merge_strategy::LoopMergeStrategy as MirroredLoopMergeStrategy;
pub use merge::once_absolute_merge_strategy::OnceAbsoluteMergeStrategy;
pub use merge::once_absolute_merge_strategy::OnceAbsoluteMergeStrategy as MirroredOnceAbsoluteMerge;
pub use metadata::collection_row_data::CollectionRowData;
pub use metadata::map_row_data::MapRowData;
pub use metadata::row_data::RowData as MirroredRowData;
pub use metadata::write_basic_parameter::WriteBasicParameter as MirroredWriteBasicParameter;
pub use metadata::write_sheet::WriteSheet as MirroredWriteSheet;
pub use metadata::write_table::WriteTable as MirroredWriteTable;
pub use metadata::write_workbook::WriteWorkbook as MirroredWriteWorkbook;
pub use property::excel_write_head_property::ExcelWriteHeadProperty;
pub use style::abstract_cell_style_strategy::AbstractCellStyleStrategy;
pub use style::abstract_vertical_cell_style_strategy::AbstractVerticalCellStyleStrategy;
pub use style::default_style::DefaultStyle;
pub use style::horizontal_cell_style_strategy::HorizontalCellStyleStrategy;
pub use style::vertical_cell_style_strategy::VerticalCellStyleStrategy;
pub use style::column::longest_match_column_width_style_strategy::LongestMatchColumnWidthStyleStrategy;
pub use style::column::simple_column_width_style_strategy::SimpleColumnWidthStyleStrategy;
pub use style::row::simple_row_height_style_strategy::SimpleRowHeightStyleStrategy;
pub use metadata::style::write_font::{
    excel_font_style_from_write_font, merge_excel_font_style, merge_write_font,
};
use metadata::style::write_cell_style::merge_write_cell_style;
use metadata::style::write_font::merge_excel_font_style as merge_handler_font_style;
use biff8::{
    date_to_excel_serial, date_to_excel_serial_with_windowing, datetime_to_excel_serial,
    datetime_to_excel_serial_with_windowing, Biff8Book, Biff8Cell, Biff8Merge, Biff8Sheet,
    Biff8StyleRequest, Biff8StyleTable, Biff8Value,
};

/// Cloneable, explicitly closeable output stream used by stateful writers.
///
/// Clones address the same underlying writer. Closing any clone drops the
/// underlying writer and makes subsequent writes fail with `BrokenPipe`, which
/// gives Rust callers an observable equivalent of Java `OutputStream.close()`.
pub struct ExcelOutputStream<W> {
    inner: Arc<Mutex<Option<W>>>,
}

impl<W> ExcelOutputStream<W> {
    /// Wraps an owned byte writer.
    #[must_use]
    pub fn new(writer: W) -> Self {
        Self {
            inner: Arc::new(Mutex::new(Some(writer))),
        }
    }

    /// Closes the shared stream, flushing it before ownership is released.
    ///
    /// # Errors
    ///
    /// Returns an error when the lock is poisoned or the final flush fails.
    pub fn close(&self) -> std::io::Result<()>
    where
        W: Write,
    {
        let mut guard = self
            .inner
            .lock()
            .map_err(|_| std::io::Error::other("output stream lock poisoned"))?;
        if let Some(mut writer) = guard.take() {
            writer.flush()?;
        }
        Ok(())
    }

    /// Returns whether the stream has been closed.
    #[must_use]
    pub fn is_closed(&self) -> bool {
        self.inner.lock().map_or(true, |writer| writer.is_none())
    }

    /// Runs a read-only callback against the underlying writer.
    ///
    /// Returns `None` after the stream is closed or if its lock was poisoned.
    pub fn with_inner<R>(&self, inspect: impl FnOnce(&W) -> R) -> Option<R> {
        self.inner
            .lock()
            .ok()
            .and_then(|writer| writer.as_ref().map(inspect))
    }

    /// Recovers the underlying writer when this is its only handle and it is open.
    ///
    /// # Errors
    ///
    /// Returns the handle when another clone exists, the stream is closed, or
    /// its lock was poisoned.
    pub fn into_inner(self) -> std::result::Result<W, Self> {
        match Arc::try_unwrap(self.inner) {
            Ok(inner) => match inner.into_inner() {
                Ok(Some(writer)) => Ok(writer),
                Ok(None) | Err(_) => Err(Self {
                    inner: Arc::new(Mutex::new(None)),
                }),
            },
            Err(inner) => Err(Self { inner }),
        }
    }
}

impl<W> Clone for ExcelOutputStream<W> {
    fn clone(&self) -> Self {
        Self {
            inner: Arc::clone(&self.inner),
        }
    }
}

impl<W> Write for ExcelOutputStream<W>
where
    W: Write,
{
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.inner
            .lock()
            .map_err(|_| std::io::Error::other("output stream lock poisoned"))?
            .as_mut()
            .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::BrokenPipe, "stream closed"))?
            .write(buffer)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner
            .lock()
            .map_err(|_| std::io::Error::other("output stream lock poisoned"))?
            .as_mut()
            .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::BrokenPipe, "stream closed"))?
            .flush()
    }
}

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
    /// Automatic trim for sheet names and string cells. (Java `autoTrim`)
    pub auto_trim: bool,
    /// Whether Excel 1904 date windowing is enabled. (Java `use1904windowing`)
    pub use_1904_windowing: bool,
    /// Locale name used for formatted output. (Java `locale`)
    pub locale: String,
    /// Whether scientific notation is used for extreme General-format numbers.
    /// (Java `useScientificFormat`)
    pub use_scientific_format: bool,
    /// Field-cache location for reflection metadata. (Java `filedCacheLocation`)
    pub filed_cache_location: CacheLocation,
    /// Whether to use a one-row constant-memory worksheet.
    pub constant_memory: bool,
    /// Whether streaming spill files use gzip (SXSSF `setCompressTempFiles`).
    ///
    /// Java mapping: `SXSSFWorkbook.setCompressTempFiles(true)` (often set in
    /// `WorkbookWriteHandler.afterWorkbookCreate`). When enabled:
    /// 1. Forces [`Self::constant_memory`] so `rust_xlsxwriter` keeps peak RAM
    ///    bounded (row window flush; avoids OOM on large batches).
    /// 2. Mirrors each data row into [`gzip_spill::GzipSheetDataWriter`] — a
    ///    true gzip tempfile (magic `1f 8b`), observable via
    ///    [`ExcelWriter::last_gzip_spill_snapshot`].
    ///
    /// **Remaining difference from POI:** POI replaces the sheet-XML spill with
    /// `GZIPSheetDataWriter` only. Here gzip is an explicit SXSSF-equivalent
    /// spill alongside the engine's constant-memory tempfile (engine tempfile
    /// stays uncompressed; final `.xlsx` is still ZIP Deflate).
    pub compress_temp_files: bool,
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
    /// Relative head row index. (Java `WriteBasicParameter.relativeHeadRowIndex`)
    pub relative_head_row_index: i32,
    /// Whether headers are auto-merged. (Java `WriteBasicParameter.automaticMergeHead`)
    pub automatic_merge_head: bool,
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
    /// Whether a stateful [`ExcelOutputStream`] is closed by `finish`.
    pub auto_close_stream: bool,
    /// Whether `finish_on_exception` emits rows accumulated before an error.
    pub write_excel_on_exception: bool,
    /// Java-style globally registered converters.
    pub converters: ConverterRegistry,
    /// Template file path. (Java `WriteWorkbook.templateFile`)
    ///
    /// When set, XLSX writes open this workbook as the write base and append
    /// typed rows after existing template content — matching Java
    /// `ExcelWriterBuilder.withTemplate(File)`. Default path preserves
    /// `styles.xml` / `mergeCells` via ZIP/OOXML; see
    /// [`Self::use_legacy_template_seed`] for the explicit value-only fallback.
    pub template_file: Option<PathBuf>,
    /// In-memory template bytes. (Java `WriteWorkbook.templateInputStream`)
    ///
    /// Builder helpers clear the other source so only one is active.
    pub template_bytes: Option<Vec<u8>>,
    /// When `true`, `with_template` uses the legacy calamine → `rust_xlsxwriter`
    /// value-replay path (styles/merges **not** preserved).
    ///
    /// Default is `false`: ZIP/OOXML preserve (`styles.xml` + `mergeCells` kept;
    /// new sheets are added as empty worksheet parts without rewriting existing
    /// sheets). Prefer leaving this off unless you explicitly need the legacy seed.
    pub use_legacy_template_seed: bool,
}

impl Default for WriteOptions {
    fn default() -> Self {
        Self {
            sheet_name: "Sheet1".to_owned(),
            sheet_index: None,
            auto_trim: true,
            use_1904_windowing: false,
            locale: "default".to_owned(),
            use_scientific_format: false,
            filed_cache_location: CacheLocation::ThreadLocal,
            constant_memory: false,
            compress_temp_files: false,
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
            auto_close_stream: true,
            write_excel_on_exception: false,
            converters: ConverterRegistry::default(),
            relative_head_row_index: 0,
            automatic_merge_head: true,
            template_file: None,
            template_bytes: None,
            use_legacy_template_seed: false,
        }
    }
}

/// Global write flags copied from [`WriteOptions`] for cell emission.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct WriteGlobalFlags {
    /// Automatic trim for sheet names and string cells.
    auto_trim: bool,
    /// Whether Excel 1904 date windowing is enabled.
    use_1904_windowing: bool,
    /// Whether scientific notation is used for extreme General-format numbers.
    use_scientific_format: bool,
}

impl From<&WriteOptions> for WriteGlobalFlags {
    fn from(options: &WriteOptions) -> Self {
        Self {
            auto_trim: options.auto_trim,
            use_1904_windowing: options.use_1904_windowing,
            use_scientific_format: options.use_scientific_format,
        }
    }
}

/// Returns the worksheet name after applying [`WriteOptions::auto_trim`].
fn effective_sheet_name(options: &WriteOptions) -> String {
    if options.auto_trim {
        options.sheet_name.trim().to_owned()
    } else {
        options.sheet_name.clone()
    }
}

/// Trims string cell text when auto-trim is enabled.
fn maybe_trim_cell_string(value: &str, auto_trim: bool) -> String {
    if auto_trim {
        value.trim().to_owned()
    } else {
        value.to_owned()
    }
}

/// Mirrors Java/reader extreme-magnitude scientific formatting threshold.
fn is_scientific_magnitude(value: f64) -> bool {
    let absolute = value.abs();
    absolute >= 1E11 || (absolute <= 1E-10 && absolute > 0.0)
}

/// Typed worksheet metadata used by [`ExcelWriter`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheet<T> {
    options: WriteOptions,
    marker: PhantomData<T>,
}

impl<T> WriteSheet<T> {
    /// Creates worksheet metadata from a complete option set.
    #[must_use]
    pub fn from_options(options: WriteOptions) -> Self {
        Self {
            options,
            marker: PhantomData,
        }
    }

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

    /// Enables compressed / disk-spill temp files for bulk writes.
    ///
    /// Java: `SXSSFWorkbook.setCompressTempFiles(bool)`. Also turns on
    /// [`Self::constant_memory`] so rows flush to disk instead of growing in RAM.
    #[must_use]
    pub const fn compress_temp_files(mut self, enabled: bool) -> Self {
        self.options.compress_temp_files = enabled;
        if enabled {
            self.options.constant_memory = true;
        }
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

/// Ensures a gzip spill writer exists for `sheet_name` when compress is on.
fn ensure_gzip_spill<'a>(
    spills: &'a mut HashMap<String, gzip_spill::GzipSheetDataWriter>,
    sheet_name: &str,
    compress: bool,
) -> Result<Option<&'a mut gzip_spill::GzipSheetDataWriter>> {
    if !compress {
        return Ok(None);
    }
    if !spills.contains_key(sheet_name) {
        spills.insert(
            sheet_name.to_owned(),
            gzip_spill::GzipSheetDataWriter::create_owned(sheet_name)?,
        );
    }
    Ok(spills.get_mut(sheet_name))
}

/// Stateful XLSX or single-sheet CSV writer matching Java `ExcelWriter`'s lifecycle.
#[allow(clippy::struct_excessive_bools)]
pub struct ExcelWriter {
    path: PathBuf,
    output_stream: Option<Box<dyn Write + Send>>,
    close_stream: Option<Box<dyn FnOnce() -> std::io::Result<()> + Send>>,
    workbook: Workbook,
    xls_book: Biff8Book,
    handlers: Vec<Box<dyn WriteHandler>>,
    sheets: HashMap<String, StatefulSheetState>,
    sheet_indexes: HashMap<usize, String>,
    csv_writer: Option<csv::Writer<CsvEncodingWriter>>,
    csv_capture: Option<CapturedOutput>,
    csv_charset: CsvCharset,
    csv_with_bom: bool,
    started: bool,
    finished: bool,
    auto_close_stream: bool,
    write_excel_on_exception: bool,
    password: Option<String>,
    converters: ConverterRegistry,
    /// Workbook-level spill preference from the builder. (Java SXSSF `setCompressTempFiles`)
    compress_temp_files: bool,
    /// Workbook-level constant-memory default from the builder.
    default_constant_memory: bool,
    template_file: Option<PathBuf>,
    template_bytes: Option<Vec<u8>>,
    /// First-write markers for sheets present in a `withTemplate` package.
    template_pending_rows: HashMap<String, u32>,
    /// ZIP/OOXML package used when preserving template styles and merges.
    template_package: Option<template_write::TemplatePackage>,
    /// OLE/BIFF8 package used when `with_template` targets a `.xls` workbook.
    ///
    /// Java mapping: `HSSFWorkbook(template)` + append cells; unmodified BIFF
    /// records are copied verbatim ([`biff8::Biff8TemplatePackage`]).
    xls_template: Option<biff8::Biff8TemplatePackage>,
    /// Explicit legacy value-replay for `with_template` (styles/merges discarded).
    use_legacy_template_seed: bool,
    /// Active gzip spill writers keyed by sheet name (when `compress_temp_files`).
    gzip_spills: HashMap<String, gzip_spill::GzipSheetDataWriter>,
    /// Last finished gzip spill snapshot (for tests / observability).
    last_gzip_spill: Option<gzip_spill::GzipSpillSnapshot>,
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
            output_stream: None,
            close_stream: None,
            workbook: Workbook::new(),
            xls_book: Biff8Book::default(),
            handlers,
            sheets: HashMap::new(),
            sheet_indexes: HashMap::new(),
            csv_writer: None,
            csv_capture: None,
            csv_charset: options.charset,
            csv_with_bom: options.with_bom,
            started: false,
            finished: false,
            auto_close_stream: options.auto_close_stream,
            write_excel_on_exception: options.write_excel_on_exception,
            password: options.password,
            converters: options.converters,
            compress_temp_files: options.compress_temp_files,
            default_constant_memory: options.constant_memory || options.compress_temp_files,
            template_file: options.template_file,
            template_bytes: options.template_bytes,
            template_pending_rows: HashMap::new(),
            template_package: None,
            xls_template: None,
            use_legacy_template_seed: options.use_legacy_template_seed,
            gzip_spills: HashMap::new(),
            last_gzip_spill: None,
        }
    }

    /// Creates a stateful writer backed by a cloneable output stream.
    #[must_use]
    pub fn with_output_stream<W>(
        logical_path: impl Into<PathBuf>,
        output: ExcelOutputStream<W>,
        handlers: Vec<Box<dyn WriteHandler>>,
        options: WriteOptions,
    ) -> Self
    where
        W: Write + Send + 'static,
    {
        let write_output = output.clone();
        let close_stream = Box::new(move || output.close());
        Self {
            path: logical_path.into(),
            output_stream: Some(Box::new(write_output)),
            close_stream: Some(close_stream),
            workbook: Workbook::new(),
            xls_book: Biff8Book::default(),
            handlers,
            sheets: HashMap::new(),
            sheet_indexes: HashMap::new(),
            csv_writer: None,
            csv_capture: None,
            csv_charset: options.charset,
            csv_with_bom: options.with_bom,
            started: false,
            finished: false,
            auto_close_stream: options.auto_close_stream,
            write_excel_on_exception: options.write_excel_on_exception,
            password: options.password,
            converters: options.converters,
            compress_temp_files: options.compress_temp_files,
            default_constant_memory: options.constant_memory || options.compress_temp_files,
            template_file: options.template_file,
            template_bytes: options.template_bytes,
            template_pending_rows: HashMap::new(),
            template_package: None,
            xls_template: None,
            use_legacy_template_seed: options.use_legacy_template_seed,
            gzip_spills: HashMap::new(),
            last_gzip_spill: None,
        }
    }

    /// Writes a batch to a worksheet, appending when the sheet was used before.
    ///
    /// XLSX and BIFF8 (`.xls`) permit multiple sheets. CSV permits repeated writes
    /// to one logical sheet, matching Java `EasyExcel`'s stateful writer.
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
        } else if is_xls_path(&self.path) {
            self.write_xls_batch::<T, I>(rows, sheet)?;
        } else {
            self.write_xlsx_batch::<T, I>(rows, sheet)?;
        }
        debug_assert!(self.resolve_sheet_name(sheet.options()).is_some());
        Ok(self)
    }

    /// Three-arg write with an explicit `WriteTable`, mirroring Java
    /// `ExcelWriter.write(Collection, WriteSheet, WriteTable)`.
    ///
    /// Phase 4 addition: this overload is the canonical entry point used
    /// when a single sheet contains multiple tables (e.g. one row block
    /// followed by a second typed block). The table options
    /// (`table_no`, `need_head`, `head_style`) override the parent
    /// sheet's options via [`crate::builder::excel_writer_table_builder::merge_table_options`].
    ///
    /// For backward compatibility this overload currently delegates to
    /// the two-arg `write` path. The merged options are applied to the
    /// sheet for the duration of this batch.
    ///
    /// # Errors
    ///
    /// Same as `write(rows, sheet)`. In addition, returns an error when
    /// the writer is finished.
    pub fn write_with_table<T, I>(
        &mut self,
        rows: I,
        sheet: &WriteSheet<T>,
        table: &MirroredWriteTable,
    ) -> Result<&mut Self>
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
        // Merge table-level options into the sheet before writing.
        // (Java: ExcelBuilderImpl.addContent(Collection, WriteSheet, WriteTable))
        let merged = crate::builder::excel_writer_table_builder::merge_table_options(
            sheet.options(),
            table,
        );
        // Rebuild a WriteSheet with merged options for this batch.
        let sheet_with_table: WriteSheet<T> = WriteSheet::from_options(merged);
        self.write(rows, &sheet_with_table)?;
        Ok(self)
    }

    /// Returns the logical output path used by Java-style builder facades.
    #[must_use]
    pub fn output_path(&self) -> &std::path::Path {
        &self.path
    }

    /// Appends raw bytes to the BIFF8 output stream. These bytes are
    /// written as an "Images" OLE stream in the CFB container when
    /// the file is serialized. Used for embedding image data in XLS.
    pub fn write_raw_bytes(&mut self, bytes: &[u8]) -> &mut Self {
        self.xls_book.write_raw_bytes(bytes);
        self
    }

    /// Encodes image bytes as BIFF8 Obj + MSODrawing + Escher BSE
    /// records (POI HSSF compatible) and embeds them in the output.
    pub fn write_image(&mut self, image_data: &[u8], col: u8, row: u32) -> &mut Self {
        self.xls_book.write_image(image_data, col, row);
        self
    }

    /// Returns whether [`WriteOptions::template_file`] / `template_bytes` is set.
    ///
    /// Mirrors Java `WriteWorkbookHolder.getTempTemplateInputStream() != null`.
    #[must_use]
    pub fn has_template_configured(&self) -> bool {
        template_write::has_template(
            self.template_file.as_deref(),
            self.template_bytes.as_deref(),
        )
    }

    /// Returns the configured template file, if any.
    #[must_use]
    pub fn template_file(&self) -> Option<&std::path::Path> {
        self.template_file.as_deref()
    }

    /// Returns the configured in-memory template bytes, if any.
    #[must_use]
    pub fn template_bytes(&self) -> Option<&[u8]> {
        self.template_bytes.as_deref()
    }

    /// Marks the writer finished without persisting workbook output.
    ///
    /// Used when a [`WriteFillExecutor`] already wrote the filled package.
    pub(crate) fn mark_finished(&mut self) {
        self.finished = true;
    }

    /// Saves and closes the writer. Repeated calls are no-ops.
    ///
    /// # Errors
    ///
    /// Returns an output or handler error.
    pub fn finish(&mut self) -> Result<()> {
        self.finish_with_exception(false)
    }

    /// Finishes after a write-side exception.
    ///
    /// By default accumulated workbook data is discarded. Set
    /// [`WriteOptions::write_excel_on_exception`] to emit it, matching Java
    /// `EasyExcel`'s `writeExcelOnException` switch.
    ///
    /// # Errors
    ///
    /// Returns an output, close, or handler error.
    pub fn finish_on_exception(&mut self) -> Result<()> {
        self.finish_with_exception(true)
    }

    fn finish_with_exception(&mut self, on_exception: bool) -> Result<()> {
        if self.finished {
            return Ok(());
        }
        self.start()?;
        if let Err(error) = self.finish_gzip_spills() {
            self.finished = true;
            return Err(error);
        }
        self.finished = true;
        let write_excel = !on_exception || self.write_excel_on_exception;
        let mut result = Ok(());
        if is_csv_path(&self.path) {
            let writer = self
                .csv_writer
                .take()
                .expect("a successfully started CSV writer must own its record writer");
            if let Err(error) = finish_csv_record_writer(writer) {
                result = Err(error);
            }
            if write_excel && let Some(capture) = self.csv_capture.take() {
                match take_captured_output(&capture).and_then(|bytes| {
                    let output = self
                        .output_stream
                        .as_mut()
                        .expect("CSV capture requires an output stream");
                    output.write_all(&bytes)?;
                    output.flush()?;
                    Ok(())
                }) {
                    Ok(()) => {}
                    Err(error) => result = Err(error),
                }
            }
        } else if write_excel && is_xls_path(&self.path) {
            let save_result = if let Some(package) = self.xls_template.take() {
                if let Some(output) = self.output_stream.as_mut() {
                    package.save_to_writer(output.as_mut())
                } else {
                    package.save_to_path(&self.path)
                }
            } else if let Some(output) = self.output_stream.as_mut() {
                self.xls_book.write_to(output.as_mut())
            } else {
                save_xls_book(&self.xls_book, &self.path)
            };
            if let Err(error) = save_result {
                result = Err(error);
            }
        } else if write_excel {
            let save_result = if let Some(package) = self.template_package.take() {
                save_template_package(
                    &package,
                    &self.path,
                    self.output_stream
                        .as_mut()
                        .map(|output| output.as_mut() as &mut (dyn Write + Send)),
                    self.password.as_deref(),
                )
            } else if let Some(output) = self.output_stream.as_mut() {
                save_workbook_to_writer(
                    &mut self.workbook,
                    output.as_mut(),
                    self.password.as_deref(),
                )
            } else {
                save_workbook(&mut self.workbook, &self.path, self.password.as_deref())
            };
            if let Err(error) = save_result {
                result = Err(error);
            }
        }
        let context = WriteWorkbookContext::new(&self.path);
        if let Err(error) = after_workbook(&mut self.handlers, &context) {
            result = Err(error);
        }
        if self.auto_close_stream
            && let Some(close) = self.close_stream.take()
            && let Err(error) = close()
        {
            result = Err(ExcelError::Io(error));
        }
        result
    }

    /// Returns whether [`Self::finish`] completed successfully.
    #[must_use]
    pub const fn is_finished(&self) -> bool {
        self.finished
    }

    /// Returns the underlying `rust_xlsxwriter` workbook for advanced XLSX customization.
    ///
    /// Callers are responsible for preserving valid worksheet names and
    /// workbook invariants. CSV writers do not use this workbook.
    #[must_use]
    pub fn workbook_mut(&mut self) -> &mut Workbook {
        &mut self.workbook
    }

    /// Enables SXSSF-style compressed / disk-spill temp files for later sheets.
    ///
    /// Java mapping: `SXSSFWorkbook.setCompressTempFiles(true)`, typically called from
    /// `WorkbookWriteHandler.afterWorkbookCreate`. Call this before the first
    /// `write` that creates a worksheet. Already-created sheets keep their mode.
    pub fn set_compress_temp_files(&mut self, enabled: bool) -> &mut Self {
        self.compress_temp_files = enabled;
        if enabled {
            self.default_constant_memory = true;
        }
        self
    }

    /// Returns whether workbook-level temp-file compression / spill is enabled.
    #[must_use]
    pub const fn compress_temp_files_enabled(&self) -> bool {
        self.compress_temp_files
    }

    /// Last finished gzip spill snapshot (Java SXSSF compressed temp observability).
    ///
    /// Populated when [`Self::finish`] closes active [`gzip_spill::GzipSheetDataWriter`]s.
    #[must_use]
    pub const fn last_gzip_spill_snapshot(&self) -> Option<&gzip_spill::GzipSpillSnapshot> {
        self.last_gzip_spill.as_ref()
    }

    /// Finishes active gzip spill writers and retains the last snapshot.
    fn finish_gzip_spills(&mut self) -> Result<()> {
        let spills = std::mem::take(&mut self.gzip_spills);
        for (_, spill) in spills {
            let reader = spill.finish()?;
            self.last_gzip_spill = Some(reader.snapshot());
        }
        Ok(())
    }

    /// Applies workbook-level spill defaults onto a sheet's write options.
    fn apply_workbook_spill_defaults(&self, options: &mut WriteOptions) {
        if self.compress_temp_files {
            options.compress_temp_files = true;
        }
        if self.default_constant_memory || options.compress_temp_files {
            options.constant_memory = true;
        }
    }

    fn start(&mut self) -> Result<()> {
        if self.started {
            return Ok(());
        }
        validate_stateful_backend(&self.path, self.password.as_deref())?;
        if template_write::has_template(
            self.template_file.as_deref(),
            self.template_bytes.as_deref(),
        ) {
            if is_csv_path(&self.path) {
                return Err(ExcelError::Unsupported(
                    "csv cannot use template.".to_owned(),
                ));
            }
            if is_xls_path(&self.path) {
                // Java: withTemplate(.xls) → HSSFWorkbook(template) + append.
                let bytes = template_write::load_template_bytes(
                    self.template_file.as_deref(),
                    self.template_bytes.as_deref(),
                )?;
                if !biff8::looks_like_xls(&bytes) {
                    return Err(ExcelError::Format(
                        "xls with_template requires an OLE .xls workbook".to_owned(),
                    ));
                }
                let package = biff8::Biff8TemplatePackage::from_bytes(&bytes)?;
                for (index, name) in package.sheet_names().into_iter().enumerate() {
                    let next_row = package.next_row_for_sheet(&name)?;
                    self.sheet_indexes.insert(index, name.clone());
                    self.template_pending_rows.insert(name, next_row);
                }
                self.xls_template = Some(package);
            } else {
                template_write::validate_template_source(
                    self.template_file.as_deref(),
                    self.template_bytes.as_deref(),
                )?;
                let bytes = template_write::load_template_bytes(
                    self.template_file.as_deref(),
                    self.template_bytes.as_deref(),
                )?;
                if self.use_legacy_template_seed {
                    // Explicit legacy fallback: value replay without styles/merges.
                    let sheets = template_write::load_template_sheets(&bytes)?;
                    template_write::seed_workbook_from_template(&mut self.workbook, &sheets)?;
                    for (index, sheet) in sheets.into_iter().enumerate() {
                        self.sheet_indexes.insert(index, sheet.name.clone());
                        self.template_pending_rows
                            .insert(sheet.name, sheet.next_row);
                    }
                } else {
                    // Default ZIP preserve path: keep styles.xml / mergeCells, append sheetData.
                    let package = template_write::TemplatePackage::from_bytes(&bytes)?;
                    for (index, name) in package.sheet_names()?.into_iter().enumerate() {
                        let next_row = package.next_row_for_sheet(&name)?;
                        self.sheet_indexes.insert(index, name.clone());
                        self.template_pending_rows.insert(name, next_row);
                    }
                    self.template_package = Some(package);
                }
            }
        }
        sort_handlers(&mut self.handlers);
        let context = WriteWorkbookContext::new(&self.path);
        before_workbook(&mut self.handlers, &context)?;
        if is_csv_path(&self.path) {
            if self.output_stream.is_some() {
                let capture = CapturedOutput::default();
                self.csv_writer = Some(create_csv_record_writer(
                    Box::new(capture.clone()),
                    &self.csv_charset,
                    self.csv_with_bom,
                )?);
                self.csv_capture = Some(capture);
            } else {
                self.csv_writer = Some(create_stateful_csv_writer(
                    &self.path,
                    &self.csv_charset,
                    self.csv_with_bom,
                )?);
            }
        }
        self.started = true;
        Ok(())
    }

    fn write_xls_batch<T, I>(&mut self, rows: I, sheet: &WriteSheet<T>) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        if self.xls_template.is_some() {
            return self.write_xls_batch_onto_template::<T, I>(rows, sheet);
        }
        let requested_name = sheet.options().sheet_name.clone();
        let existing_name = self.resolve_sheet_name(sheet.options());
        let sheet_name = existing_name.unwrap_or_else(|| requested_name.clone());
        let (state, is_new) = if let Some(state) = self.sheets.get(&sheet_name).cloned() {
            validate_stateful_schema(&sheet_name, &state, T::schema())?;
            (state, false)
        } else {
            let mut options = sheet.options().clone();
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
        let progress = {
            if is_new {
                let biff_sheet = self.xls_book.sheet_mut(&sheet_name);
                biff_sheet.next_row = state.next_row;
                biff_sheet.next_data_index = state.next_data_index;
            }
            append_rows_to_biff8_sheet::<T, I>(
                &mut self.xls_book,
                &sheet_name,
                &state.options,
                rows,
                &mut self.handlers,
                WriteProgress {
                    next_row: state.next_row,
                    next_data_index: state.next_data_index,
                },
                is_new,
            )?
        };
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
        self.remember_sheet_index(sheet.options().sheet_index, &sheet_name);
        Ok(())
    }

    /// Appends typed rows onto a record-preserving `.xls` template package.
    ///
    /// Mirrors [`Self::write_xlsx_batch_onto_template_package`] for HSSF/BIFF8.
    /// Creating sheets absent from the template remains unsupported (MVP).
    fn write_xls_batch_onto_template<T, I>(
        &mut self,
        rows: I,
        sheet: &WriteSheet<T>,
    ) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let sheet_names = {
            let package = self
                .xls_template
                .as_ref()
                .expect("xls template must exist for BIFF preserve path");
            package.sheet_names()
        };
        let (_target_index, target_name, create_new) = template_write::resolve_package_target(
            &sheet_names,
            sheet.options().sheet_index,
            &sheet.options().sheet_name,
        );
        if create_new {
            return Err(ExcelError::Unsupported(
                "xls template cannot create sheets absent from the template".to_owned(),
            ));
        }
        let sheet_name = target_name;
        let (state, is_new) = if let Some(state) = self.sheets.get(&sheet_name).cloned() {
            validate_stateful_schema(&sheet_name, &state, T::schema())?;
            (state, false)
        } else {
            let mut options = sheet.options().clone();
            options.sheet_name = sheet_name.clone();
            options.converters = self.converters.merged_with(&options.converters);
            let next_row = self
                .template_pending_rows
                .get(&sheet_name)
                .copied()
                .unwrap_or(0);
            (
                StatefulSheetState {
                    schema: T::schema(),
                    metadata: *T::write_metadata(),
                    options,
                    next_row,
                    next_data_index: 0,
                },
                true,
            )
        };
        let sheet_context = WriteSheetContext::new(&sheet_name);
        if is_new {
            before_sheet(&mut self.handlers, &sheet_context)?;
        }
        let first_write = self.template_pending_rows.remove(&sheet_name).is_some() || is_new;
        let write_head = first_write;
        let append_rows = collect_template_append_rows::<T, I>(
            &state.options,
            rows,
            write_head,
            state.next_data_index,
        )?;
        let next_row = {
            let package = self
                .xls_template
                .as_mut()
                .expect("xls template must exist for BIFF preserve path");
            package.append_rows(&sheet_name, &append_rows)?
        };
        let head_rows = if write_head {
            usize::try_from(head_rows_for_schema(T::schema(), &state.options)?).unwrap_or(0)
        } else {
            0
        };
        let data_added = append_rows.len().saturating_sub(head_rows);
        if is_new {
            after_sheet(&mut self.handlers, &sheet_context)?;
        }
        self.sheets.insert(
            sheet_name.clone(),
            StatefulSheetState {
                next_row,
                next_data_index: state.next_data_index.saturating_add(data_added),
                ..state
            },
        );
        self.remember_sheet_index(sheet.options().sheet_index, &sheet_name);
        Ok(())
    }

    fn write_xlsx_batch<T, I>(&mut self, rows: I, sheet: &WriteSheet<T>) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let requested_name = sheet.options().sheet_name.clone();
        if self.template_package.is_some() {
            return self.write_xlsx_batch_onto_template_package::<T, I>(rows, sheet);
        }
        if let Some(sheet_name) = self.resolve_sheet_name(sheet.options()) {
            if let Some(start_row) = self.template_pending_rows.remove(&sheet_name) {
                let mut options = sheet.options().clone();
                options.converters = self.converters.merged_with(&options.converters);
                self.apply_workbook_spill_defaults(&mut options);
                // Preserve the real template sheet name (index-based Java `.sheet()`).
                options.sheet_name = sheet_name.clone();
                let worksheet = self
                    .workbook
                    .worksheet_from_name(&sheet_name)
                    .map_err(format_error)?;
                let sheet_context = WriteSheetContext::new(&sheet_name);
                before_sheet(&mut self.handlers, &sheet_context)?;
                let compress = options.compress_temp_files;
                let progress = {
                    let spill =
                        ensure_gzip_spill(&mut self.gzip_spills, &sheet_name, compress)?;
                    append_rows_to_worksheet_with_gzip::<T, I>(
                        worksheet,
                        &options,
                        rows,
                        &mut self.handlers,
                        WriteProgress {
                            next_row: start_row,
                            next_data_index: 0,
                        },
                        true,
                        T::write_metadata(),
                        spill,
                    )?
                };
                after_sheet(&mut self.handlers, &sheet_context)?;
                // Java LongestMatchColumnWidthStyleStrategy setColumnWidth after cells
                apply_handler_column_widths::<T>(worksheet, &options, &self.handlers)?;
                if options.auto_width || handlers_request_auto_width(&self.handlers) {
                    worksheet.autofit();
                }
                self.sheets.insert(
                    sheet_name.clone(),
                    StatefulSheetState {
                        schema: T::schema(),
                        metadata: *T::write_metadata(),
                        options,
                        next_row: progress.next_row,
                        next_data_index: progress.next_data_index,
                    },
                );
                self.remember_sheet_index(sheet.options().sheet_index, &sheet_name);
                return Ok(());
            }
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
            let compress = state.options.compress_temp_files;
            let progress = {
                let spill = ensure_gzip_spill(&mut self.gzip_spills, &sheet_name, compress)?;
                append_rows_to_worksheet_with_gzip::<T, I>(
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
                    spill,
                )?
            };
            if state.options.auto_width || handlers_request_auto_width(&self.handlers) {
                worksheet.autofit();
            }
            // Re-apply measured LongestMatch widths after incremental append.
            apply_handler_column_widths::<T>(worksheet, &state.options, &self.handlers)?;
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
        self.apply_workbook_spill_defaults(&mut options);
        let sheet_name = options.sheet_name.clone();
        let compress = options.compress_temp_files;
        let progress = {
            let spill = ensure_gzip_spill(&mut self.gzip_spills, &sheet_name, compress)?;
            write_sheet_to_workbook_with_gzip::<T, I>(
                &mut self.workbook,
                &options,
                rows,
                &mut self.handlers,
                spill,
            )?
        };
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

    /// Appends typed rows onto a ZIP-preserved template package.
    ///
    /// Keeps `styles.xml` and `mergeCells` from the template; only `sheetData`
    /// grows. When the requested sheet is absent, a new empty worksheet part is
    /// created without rewriting existing sheets.
    fn write_xlsx_batch_onto_template_package<T, I>(
        &mut self,
        rows: I,
        sheet: &WriteSheet<T>,
    ) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let sheet_names = {
            let package = self
                .template_package
                .as_ref()
                .expect("template package must exist for ZIP preserve path");
            package.sheet_names()?
        };
        let (_target_index, target_name, create_new) = template_write::resolve_package_target(
            &sheet_names,
            sheet.options().sheet_index,
            &sheet.options().sheet_name,
        );
        let sheet_name = if let Some(resolved) = self.resolve_sheet_name(sheet.options()) {
            resolved
        } else if create_new {
            let package = self
                .template_package
                .as_mut()
                .expect("template package must exist for ZIP preserve path");
            package.ensure_sheet(&target_name)?;
            self.template_pending_rows.insert(target_name.clone(), 0);
            target_name
        } else {
            target_name
        };
        let first_write = self.template_pending_rows.remove(&sheet_name).is_some()
            || !self.sheets.contains_key(&sheet_name);
        let mut options = if let Some(state) = self.sheets.get(&sheet_name) {
            validate_stateful_schema(&sheet_name, state, T::schema())?;
            state.options.clone()
        } else {
            let mut options = sheet.options().clone();
            options.converters = self.converters.merged_with(&options.converters);
            self.apply_workbook_spill_defaults(&mut options);
            options.sheet_name = sheet_name.clone();
            options
        };
        options.sheet_name = sheet_name.clone();

        let write_head = first_write;
        let next_data_index = self
            .sheets
            .get(&sheet_name)
            .map(|state| state.next_data_index)
            .unwrap_or(0);
        let append_rows =
            collect_template_append_rows::<T, I>(&options, rows, write_head, next_data_index)?;
        let sheet_context = WriteSheetContext::new(&sheet_name);
        if first_write {
            before_sheet(&mut self.handlers, &sheet_context)?;
        }
        let next_row = {
            let package = self
                .template_package
                .as_mut()
                .expect("template package must exist for ZIP preserve path");
            package.append_rows(&sheet_name, &append_rows)?
        };
        if first_write {
            after_sheet(&mut self.handlers, &sheet_context)?;
        }
        let added = append_rows.len();
        let head_rows = if write_head {
            usize::try_from(head_rows_for_schema(T::schema(), &options)?).unwrap_or(0)
        } else {
            0
        };
        let data_added = added.saturating_sub(head_rows);
        self.sheets.insert(
            sheet_name.clone(),
            StatefulSheetState {
                schema: T::schema(),
                metadata: *T::write_metadata(),
                options,
                next_row,
                next_data_index: next_data_index.saturating_add(data_added),
            },
        );
        self.remember_sheet_index(sheet.options().sheet_index, &sheet_name);
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
            .or_else(|| {
                self.template_pending_rows
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
        // XLS password is now supported via BIFF8 RC4 (Phase 5.3)
        _ => Ok(()),
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

fn uses_constant_memory_spill(options: &WriteOptions) -> bool {
    options.constant_memory || options.compress_temp_files
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

/// Writes typed rows to a new BIFF8 (`.xls`) file.
///
/// Java mapping: `EasyExcel.write(path, head).excelType(XLS).sheet().doWrite(data)`.
///
/// # Errors
///
/// Returns a conversion, worksheet-configuration, BIFF8-format, or I/O error.
pub fn write_xls<T, I>(path: &Path, options: &WriteOptions, rows: I) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    write_xls_with_handlers(path, options, rows, &mut [])
}

/// Writes typed rows to a BIFF8 file while invoking ordered write handlers.
///
/// When [`WriteOptions`] carries a template, uses
/// [`biff8::Biff8TemplatePackage`] (Java `withTemplate` + `doWrite` on HSSF).
/// Password protection remains [`ExcelError::Unsupported`].
///
/// # Errors
///
/// Returns a conversion, handler, BIFF8-format, template, or I/O error.
pub fn write_xls_with_handlers<T, I>(
    path: &Path,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    validate_xls_options(options)?;
    sort_handlers(handlers);
    let workbook_context = WriteWorkbookContext::new(path);
    before_workbook(handlers, &workbook_context)?;

    if template_write::has_template(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    ) {
        write_xls_onto_template::<T, I>(path, None, options, rows, handlers)?;
        after_workbook(handlers, &workbook_context)?;
        return Ok(());
    }

    let mut book = Biff8Book::default();
    write_sheet_to_biff8_book::<T, I>(&mut book, options, rows, handlers)?;
    // Phase 5.3: BIFF8 RC4 encryption
    if let Some(password) = &options.password {
        let raw_bytes = book.to_cfb_bytes()?;
        let (encrypted, _salt, _vh) =
            crate::biff8::encrypt::encrypt_biff8_stream(&raw_bytes, password);
        std::fs::write(path, &encrypted).map_err(ExcelError::from)?;
    } else {
        save_xls_book(&book, path)?;
    }
    after_workbook(handlers, &workbook_context)?;
    Ok(())
}

/// Writes typed rows as BIFF8 bytes to an arbitrary writer.
///
/// # Errors
///
/// Returns a conversion, handler, BIFF8-format, or stream I/O error.
pub fn write_xls_to_writer<T, I, W>(
    logical_path: &Path,
    mut output: W,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
    W: Write + Send,
{
    validate_xls_options(options)?;
    sort_handlers(handlers);
    let workbook_context = WriteWorkbookContext::new(logical_path);
    before_workbook(handlers, &workbook_context)?;

    if template_write::has_template(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    ) {
        write_xls_onto_template::<T, I>(
            logical_path,
            Some(&mut output),
            options,
            rows,
            handlers,
        )?;
        after_workbook(handlers, &workbook_context)?;
        return Ok(());
    }

    let mut book = Biff8Book::default();
    write_sheet_to_biff8_book::<T, I>(&mut book, options, rows, handlers)?;
    book.write_to(&mut output)?;
    output.flush()?;
    after_workbook(handlers, &workbook_context)?;
    Ok(())
}

fn validate_xls_options(_options: &WriteOptions) -> Result<()> {
    // XLS password is now supported via BIFF8 RC4 (Phase 5.3)
    Ok(())
}

/// Writes typed rows onto an existing `.xls` template (Java `withTemplate` + `doWrite`).
///
/// Uses [`biff8::Biff8TemplatePackage`] so unmodified BIFF records stay intact.
fn write_xls_onto_template<T, I>(
    path: &Path,
    output: Option<&mut (dyn Write + Send)>,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    validate_xls_options(options)?;
    let bytes = template_write::load_template_bytes(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    )?;
    if !biff8::looks_like_xls(&bytes) {
        return Err(ExcelError::Format(
            "xls with_template requires an OLE .xls workbook".to_owned(),
        ));
    }
    let mut package = biff8::Biff8TemplatePackage::from_bytes(&bytes)?;
    let sheet_names = package.sheet_names();
    let (_target_index, target_name, create_new) = template_write::resolve_package_target(
        &sheet_names,
        options.sheet_index,
        &options.sheet_name,
    );
    if create_new {
        return Err(ExcelError::Unsupported(
            "xls template cannot create sheets absent from the template".to_owned(),
        ));
    }
    let mut write_options = options.clone();
    write_options.sheet_name = target_name.clone();
    let append_rows = collect_template_append_rows::<T, I>(&write_options, rows, true, 0)?;
    let sheet_context = WriteSheetContext::new(&target_name);
    before_sheet(handlers, &sheet_context)?;
    package.append_rows(&target_name, &append_rows)?;
    after_sheet(handlers, &sheet_context)?;
    match output {
        Some(writer) => package.save_to_writer(writer),
        None => package.save_to_path(path),
    }
}

fn save_xls_book(book: &Biff8Book, path: &Path) -> Result<()> {
    if let Some(parent) = path.parent()
        && !parent.as_os_str().is_empty()
    {
        std::fs::create_dir_all(parent)?;
    }
    let mut file = File::create(path)?;
    book.write_to(&mut file)?;
    file.flush()?;
    Ok(())
}

fn write_sheet_to_biff8_book<T, I>(
    book: &mut Biff8Book,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let sheet_name = effective_sheet_name(options);
    let mut write_options = options.clone();
    write_options.sheet_name = sheet_name.clone();
    book.use_1904_windowing = write_options.use_1904_windowing;
    let sheet_context = WriteSheetContext::new(&sheet_name);
    before_sheet(handlers, &sheet_context)?;
    let progress = append_rows_to_biff8_sheet::<T, I>(
        book,
        &sheet_name,
        &write_options,
        rows,
        handlers,
        WriteProgress {
            next_row: 0,
            next_data_index: 0,
        },
        true,
    )?;
    after_sheet(handlers, &sheet_context)?;
    Ok(progress)
}

/// Appends typed rows onto a BIFF8 sheet buffer (header + data + styles/merges).
///
/// Consumes [`WriteOptions`] column widths / styles / merge ranges and annotation
/// metadata (`@ColumnWidth`, `@HeadRowHeight`, `@ContentRowHeight`, `@HeadStyle`,
/// `@ContentStyle`, `@OnceAbsoluteMerge`, `@ContentLoopMerge`) — Java HSSF parity
/// for the Minimal BIFF8 subset.
fn append_rows_to_biff8_sheet<T, I>(
    book: &mut Biff8Book,
    sheet_name: &str,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
    progress: WriteProgress,
    write_head: bool,
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let WriteProgress {
        next_row: mut row_index,
        next_data_index: mut data_index,
    } = progress;
    let global = WriteGlobalFlags::from(options);
    let columns = selected_columns(T::schema(), options);
    let metadata = T::write_metadata();
    let head_rows = head_rows_for_schema(T::schema(), options)?;
    let annotation_loop_merges = annotation_loop_merges_from_columns(&columns)?;

    if write_head {
        apply_biff8_column_widths::<T>(book.sheet_mut(sheet_name), options, handlers)?;
        apply_biff8_once_absolute_merges::<T>(book.sheet_mut(sheet_name), handlers, row_index)?;
        for range in &options.merge_ranges {
            add_biff8_merge_range(book.sheet_mut(sheet_name), *range)?;
        }
    }

    if write_head && head_rows > 0 {
        write_biff8_headers(
            book,
            sheet_name,
            &columns,
            options,
            metadata,
            handlers,
            row_index,
        )?;
        // Annotation `@HeadRowHeight` / `SimpleRowHeightStyleStrategy`
        let head_height = metadata
            .head_row_height
            .or_else(|| collect_handler_head_row_height(handlers));
        if let Some(height) = head_height {
            let sheet = book.sheet_mut(sheet_name);
            for head_row in row_index..row_index + head_rows {
                let row = u16::try_from(head_row).map_err(|_| {
                    ExcelError::Format("BIFF8 row overflow".to_owned())
                })?;
                sheet.set_row_height(row, height);
            }
        }
        if options.automatic_merge_head {
            if let Some(head) = &options.dynamic_head {
                merge_biff8_dynamic_head_groups(
                    book.sheet_mut(sheet_name),
                    &columns,
                    head,
                    row_index,
                )?;
            }
        }
        row_index = row_index
            .checked_add(head_rows)
            .ok_or_else(|| ExcelError::Format("BIFF8 row overflow".to_owned()))?;
    }

    let row_list: Vec<T> = rows.into_iter().collect();
    for row in row_list {
        let content_height = metadata
            .content_row_height
            .or_else(|| collect_handler_content_row_height(handlers));
        if let Some(height) = content_height {
            let row_u16 = u16::try_from(row_index).map_err(|_| {
                ExcelError::Format("BIFF8 row overflow".to_owned())
            })?;
            book.sheet_mut(sheet_name).set_row_height(row_u16, height);
        }
        let cells = row.to_row_with_converters(&options.converters)?;
        let dynamic_columns = dynamic_columns_for_row(T::schema().is_empty(), cells.len(), options);
        let row_columns = dynamic_columns.as_deref().unwrap_or(&columns);
        let explicit_style = (!options.content_styles.is_empty())
            .then(|| &options.content_styles[data_index % options.content_styles.len()]);
        apply_biff8_loop_merges(
            book.sheet_mut(sheet_name),
            row_index,
            data_index,
            &options.loop_merges,
        )?;
        apply_biff8_loop_merges(
            book.sheet_mut(sheet_name),
            row_index,
            data_index,
            &annotation_loop_merges,
        )?;
        let row_context = WriteRowContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            is_head: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_row(&row_context)?;
        }
        for (physical_index, schema_index, column) in row_columns {
            let value = cells.get(*schema_index).unwrap_or(&CellValue::Empty);
            let mut context = WriteCellContext {
                sheet_name: sheet_name.to_owned(),
                row_index,
                column_index: to_column(*physical_index)?,
                field: (!column.field.is_empty()).then_some(column.field),
                is_head: false,
                relative_row_index: Some(data_index),
                value: value.clone(),
                skip: false,
            };
            for handler in handlers.iter_mut() {
                handler.before_cell(&mut context)?;
            }
            if !context.skip {
                let handler_style = collect_handler_cell_style(handlers, &context);
                let style_ctx = SheetStyleContext::content(explicit_style, metadata, global);
                let format_ctx = style_ctx.column(column).with_handler_cell(handler_style);
                let cell = cell_value_to_biff8_styled(
                    &context.value,
                    &mut book.styles,
                    format_ctx,
                )?;
                book.sheet_mut(sheet_name)
                    .set(row_index, *physical_index, cell)?;
            }
            for handler in handlers.iter_mut() {
                handler.after_cell(&context)?;
            }
        }
        for handler in handlers.iter_mut() {
            handler.after_row(&row_context)?;
        }
        row_index = row_index
            .checked_add(1)
            .ok_or_else(|| ExcelError::Format("BIFF8 row overflow".to_owned()))?;
        data_index += 1;
    }
    // LongestMatch / strategy widths may update after cells (Java afterCellDispose).
    apply_biff8_handler_column_widths::<T>(book.sheet_mut(sheet_name), options, handlers)?;
    let sheet = book.sheet_mut(sheet_name);
    sheet.next_row = row_index;
    sheet.next_data_index = data_index;
    Ok(WriteProgress {
        next_row: row_index,
        next_data_index: data_index,
    })
}

fn write_biff8_headers(
    book: &mut Biff8Book,
    sheet_name: &str,
    columns: &[(usize, usize, &'static ExcelColumn)],
    options: &WriteOptions,
    metadata: &ExcelWriteMetadata,
    handlers: &mut [Box<dyn WriteHandler>],
    start_row: u32,
) -> Result<()> {
    let global = WriteGlobalFlags::from(options);
    let style_ctx = SheetStyleContext::head(&options.head_style, metadata, global);
    if let Some(head) = &options.dynamic_head {
        let levels = head.iter().map(Vec::len).max().unwrap_or(0);
        for level in 0..levels {
            let row = start_row
                .checked_add(u32::try_from(level).map_err(|_| {
                    ExcelError::Format("dynamic head is too deep".to_owned())
                })?)
                .ok_or_else(|| ExcelError::Format("BIFF8 row overflow".to_owned()))?;
            for (col_idx, path) in head.iter().enumerate() {
                let label = path.get(level).cloned().unwrap_or_default();
                let (physical, column) = match columns.get(col_idx) {
                    Some((physical, _, column)) => (*physical, *column),
                    None => continue,
                };
                write_biff8_styled_text_cell(
                    book,
                    sheet_name,
                    row,
                    physical,
                    label,
                    style_ctx.column(column),
                    handlers,
                    true,
                )?;
            }
        }
    } else {
        for (physical_index, _, column) in columns {
            write_biff8_styled_text_cell(
                book,
                sheet_name,
                start_row,
                *physical_index,
                column.name.to_owned(),
                style_ctx.column(column),
                handlers,
                true,
            )?;
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn write_biff8_styled_text_cell(
    book: &mut Biff8Book,
    sheet_name: &str,
    row_index: u32,
    physical_index: usize,
    label: String,
    format_ctx: CellFormatContext<'_>,
    handlers: &mut [Box<dyn WriteHandler>],
    is_head: bool,
) -> Result<()> {
    let column_index = to_column(physical_index)?;
    let mut context = WriteCellContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        column_index,
        field: None,
        is_head,
        relative_row_index: None,
        value: CellValue::String(label.clone()),
        skip: false,
    };
    for handler in handlers.iter_mut() {
        handler.before_cell(&mut context)?;
    }
    if !context.skip {
        let handler_style = collect_handler_cell_style(handlers, &context);
        let format_ctx = format_ctx.with_handler_cell(handler_style);
        let cell =
            cell_value_to_biff8_styled(&CellValue::String(label), &mut book.styles, format_ctx)?;
        book.sheet_mut(sheet_name)
            .set(row_index, physical_index, cell)?;
    }
    for handler in handlers.iter_mut() {
        handler.after_cell(&context)?;
    }
    Ok(())
}

fn cell_value_to_biff8(value: &CellValue, global: WriteGlobalFlags) -> Result<Biff8Cell> {
    match value {
        CellValue::Empty => Ok(Biff8Cell::general(Biff8Value::Blank)),
        CellValue::String(text) | CellValue::Error(text) | CellValue::Formula(text) => {
            Ok(Biff8Cell::general(Biff8Value::Text(maybe_trim_cell_string(
                text,
                global.auto_trim,
            ))))
        }
        CellValue::Bool(flag) => Ok(Biff8Cell::general(Biff8Value::Bool(*flag))),
        CellValue::Int(number) => {
            #[allow(clippy::cast_precision_loss)]
            Ok(Biff8Cell::general(Biff8Value::Number(*number as f64)))
        }
        CellValue::Float(number) => Ok(Biff8Cell::general(Biff8Value::Number(*number))),
        CellValue::Decimal(number) => {
            let number = number
                .to_f64()
                .filter(|value| value.is_finite())
                .ok_or_else(|| {
                    ExcelError::Format("decimal value exceeds BIFF8 numeric range".to_owned())
                })?;
            Ok(Biff8Cell::general(Biff8Value::Number(number)))
        }
        CellValue::Date(date) => Ok(Biff8Cell::date_serial(date_to_excel_serial_with_windowing(
            *date,
            global.use_1904_windowing,
        ))),
        CellValue::DateTime(date_time) => Ok(Biff8Cell::datetime_serial(
            datetime_to_excel_serial_with_windowing(*date_time, global.use_1904_windowing),
        )),
        CellValue::Hyperlink { text, .. } => Ok(Biff8Cell::general(Biff8Value::Text(
            maybe_trim_cell_string(text, global.auto_trim),
        ))),
        CellValue::Comment { value, .. } => cell_value_to_biff8(value, global),
        CellValue::Images { value, images } => {
            // Write the base value; image bytes are persisted via
            // write_raw_bytes on the Biff8Book (called by caller).
            for img in images.iter() {
                let _ = img.image();
            }
            cell_value_to_biff8(value, global)
        }
        CellValue::RichText(rich) => Ok(Biff8Cell::general(Biff8Value::Text(
            maybe_trim_cell_string(rich.text_string(), global.auto_trim),
        ))),
        CellValue::Image(bytes) => {
            // Write base value, image bytes handled by caller
            let _ = bytes;
            Ok(Biff8Cell::general(Biff8Value::Blank))
        },
    }
}

/// Converts a cell value and applies FONT/XF from annotation + handler styles.
fn cell_value_to_biff8_styled(
    value: &CellValue,
    styles: &mut Biff8StyleTable,
    format_ctx: CellFormatContext<'_>,
) -> Result<Biff8Cell> {
    let cell = cell_value_to_biff8(value, format_ctx.global)?;
    let request = biff8_style_request(styles, format_ctx);
    let xf = styles.resolve_xf(&request, cell.xf);
    Ok(cell.with_xf(xf))
}

/// Builds a BIFF8 style request from the same merge order as XLSX `cell_format`.
fn biff8_style_request(
    styles: &mut Biff8StyleTable,
    context: CellFormatContext<'_>,
) -> Biff8StyleRequest {
    let mut request = Biff8StyleRequest::default();
    let mut annotation_cell = context.cell;
    if let Some(handler_style) = context.handler_cell {
        annotation_cell = Some(merge_write_cell_style(
            &handler_style,
            annotation_cell.unwrap_or_default(),
        ));
    }
    let mut font = context.font;
    if let Some(style) = annotation_cell {
        if let Some(style_font) = style.font {
            font = Some(match font {
                Some(target) => merge_handler_font_style(&style_font, target),
                None => style_font,
            });
        }
        // Remap RGB fills through the palette allocator before applying.
        let mut style = style;
        if let Some(ExcelColor::Rgb(rgb)) = style.fill_foreground_color {
            style.fill_foreground_color = Some(ExcelColor::Indexed(
                u8::try_from(styles.alloc_rgb_icv(rgb)).unwrap_or(8),
            ));
        }
        if let Some(ExcelColor::Rgb(rgb)) = style.fill_background_color {
            style.fill_background_color = Some(ExcelColor::Indexed(
                u8::try_from(styles.alloc_rgb_icv(rgb)).unwrap_or(8),
            ));
        }
        request.apply_excel_cell_style(style);
    }
    if let Some(font) = font {
        let mut font = font;
        if let Some(ExcelColor::Rgb(rgb)) = font.color {
            font.color = Some(ExcelColor::Indexed(
                u8::try_from(styles.alloc_rgb_icv(rgb)).unwrap_or(8),
            ));
        }
        request.apply_excel_font_style(font);
    }
    if let Some(style) = context.explicit {
        apply_writer_cell_style_to_request(&mut request, styles, style);
    }
    request
}

/// Maps [`CellStyle`] builder fields onto a BIFF8 style request.
fn apply_writer_cell_style_to_request(
    request: &mut Biff8StyleRequest,
    styles: &mut Biff8StyleTable,
    style: &CellStyle,
) {
    if style.bold {
        request.bold = true;
    }
    if style.italic {
        request.italic = true;
    }
    if let Some(color) = style.font_color {
        request.font_color_icv = Some(styles.alloc_rgb_icv(color));
    }
    if let Some(color) = style.background_color {
        request.fill_pattern = Some(1);
        request.fill_fg_icv = Some(styles.alloc_rgb_icv(color));
        request.fill_bg_icv = Some(64); // automatic pattern background
    }
    if let Some(alignment) = style.horizontal_alignment {
        request.halign = Some(biff8_halign(alignment));
    }
    if let Some(alignment) = style.vertical_alignment {
        request.valign = Some(biff8_valign(alignment));
    }
    if style.wrap_text {
        request.wrap = true;
    }
}

const fn biff8_halign(align: HorizontalAlignment) -> u8 {
    match align {
        HorizontalAlignment::General => 0,
        HorizontalAlignment::Left => 1,
        HorizontalAlignment::Center => 2,
        HorizontalAlignment::Right => 3,
        HorizontalAlignment::Fill => 4,
        HorizontalAlignment::Justify => 5,
        HorizontalAlignment::CenterAcross => 6,
    }
}

const fn biff8_valign(align: VerticalAlignment) -> u8 {
    match align {
        VerticalAlignment::Top => 0,
        VerticalAlignment::Center => 1,
        VerticalAlignment::Bottom => 2,
        VerticalAlignment::Justify => 3,
        VerticalAlignment::Distributed => 4,
    }
}

/// Applies explicit + annotation + handler column widths to a BIFF8 sheet.
fn apply_biff8_column_widths<T>(
    sheet: &mut Biff8Sheet,
    options: &WriteOptions,
    handlers: &[Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
{
    for (column, width) in &options.column_widths {
        let col = u8::try_from(*column).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 256 columns".to_owned())
        })?;
        sheet.set_column_width(col, *width);
    }
    let type_width = T::write_metadata().column_width;
    for (physical_index, _, column) in selected_columns(T::schema(), options) {
        let col = u8::try_from(physical_index).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 256 columns".to_owned())
        })?;
        if sheet.column_widths.contains_key(&col) {
            continue;
        }
        if let Some(width) = column.column_width.or(type_width) {
            sheet.set_column_width(col, width);
        }
    }
    apply_biff8_handler_column_widths::<T>(sheet, options, handlers)
}

/// Applies registered column-width strategies (Java `SimpleColumnWidthStyleStrategy`).
fn apply_biff8_handler_column_widths<T>(
    sheet: &mut Biff8Sheet,
    options: &WriteOptions,
    handlers: &[Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
{
    for (physical_index, _, _) in selected_columns(T::schema(), options) {
        let col = u8::try_from(physical_index).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 256 columns".to_owned())
        })?;
        if options
            .column_widths
            .iter()
            .any(|(explicit, _)| usize::from(*explicit) == physical_index)
        {
            continue;
        }
        for handler in handlers {
            if let Some(width) = handler.style_column_width(physical_index) {
                sheet.set_column_width(col, width);
            }
        }
    }
    Ok(())
}

/// Applies `@OnceAbsoluteMerge` + registered once-absolute strategies.
fn apply_biff8_once_absolute_merges<T>(
    sheet: &mut Biff8Sheet,
    handlers: &[Box<dyn WriteHandler>],
    row_offset: u32,
) -> Result<()>
where
    T: ExcelRow,
{
    if let Some(merge) = T::write_metadata().once_absolute_merge {
        apply_biff8_once_absolute_merge_property(sheet, merge, row_offset)?;
    }
    for handler in handlers {
        if let Some(merge) = handler.style_once_absolute_merge() {
            apply_biff8_once_absolute_merge_property(sheet, merge, row_offset)?;
        }
    }
    Ok(())
}

fn apply_biff8_once_absolute_merge_property(
    sheet: &mut Biff8Sheet,
    merge: easyexcel_core::OnceAbsoluteMergeProperty,
    row_offset: u32,
) -> Result<()> {
    if merge.first_row_index < 0
        || merge.last_row_index < 0
        || merge.first_column_index < 0
        || merge.last_column_index < 0
    {
        return Ok(());
    }
    #[allow(clippy::cast_sign_loss, clippy::cast_possible_truncation)]
    add_biff8_merge_range(
        sheet,
        MergeRange::new(
            (merge.first_row_index as u32).saturating_add(row_offset),
            (merge.last_row_index as u32).saturating_add(row_offset),
            merge.first_column_index as u16,
            merge.last_column_index as u16,
        ),
    )
}

fn add_biff8_merge_range(sheet: &mut Biff8Sheet, range: MergeRange) -> Result<()> {
    let first_row = u16::try_from(range.first_row).map_err(|_| {
        ExcelError::Format("BIFF8 merge row exceeds 65536".to_owned())
    })?;
    let last_row = u16::try_from(range.last_row).map_err(|_| {
        ExcelError::Format("BIFF8 merge row exceeds 65536".to_owned())
    })?;
    let first_col = u8::try_from(range.first_column).map_err(|_| {
        ExcelError::Format("BIFF8 merge column exceeds 256".to_owned())
    })?;
    let last_col = u8::try_from(range.last_column).map_err(|_| {
        ExcelError::Format("BIFF8 merge column exceeds 256".to_owned())
    })?;
    sheet.add_merge(Biff8Merge {
        first_row,
        last_row,
        first_col,
        last_col,
    })
}

fn apply_biff8_loop_merges(
    sheet: &mut Biff8Sheet,
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
        add_biff8_merge_range(
            sheet,
            MergeRange::new(row_index, last_row, strategy.column_index, last_column),
        )?;
    }
    Ok(())
}

fn merge_biff8_dynamic_head_groups(
    sheet: &mut Biff8Sheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    head: &[Vec<String>],
    start_row: u32,
) -> Result<()> {
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    for level in 0..levels {
        #[allow(clippy::cast_possible_truncation)]
        let row_index = start_row.saturating_add(level as u32);
        let mut start = 0;
        while start < head.len() {
            let mut end = start;
            while end + 1 < head.len()
                && columns.get(end).zip(columns.get(end + 1)).is_some_and(
                    |((left, _, _), (right, _, _))| left.checked_add(1) == Some(*right),
                )
                && same_dynamic_head_group(head, start, end + 1, level)
            {
                end += 1;
            }
            let label = head[start].get(level).map_or("", String::as_str);
            if end > start && !label.is_empty() {
                add_biff8_merge_range(
                    sheet,
                    MergeRange::new(
                        row_index,
                        row_index,
                        to_column(columns[start].0)?,
                        to_column(columns[end].0)?,
                    ),
                )?;
            }
            start = end + 1;
        }
    }
    Ok(())
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

    if template_write::has_template(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    ) {
        write_xlsx_onto_template_package::<T, I>(path, None, options, rows, handlers)?;
    } else {
        let mut workbook = Workbook::new();
        write_sheet_to_workbook::<T, I>(&mut workbook, options, rows, handlers)?;
        save_workbook(&mut workbook, path, options.password.as_deref())?;
    }
    after_workbook(handlers, &workbook_context)?;
    Ok(())
}

/// Writes typed rows to an arbitrary XLSX byte stream.
///
/// `logical_path` is used only by write-handler contexts. Unlike the path
/// entry point this function writes the OOXML package to `output` itself, so
/// it is suitable for HTTP response bodies and in-memory buffers.
///
/// # Errors
///
/// Returns a conversion, handler, worksheet-configuration, XLSX-format,
/// encryption, or stream I/O error.
pub fn write_xlsx_to_writer<T, I, W>(
    logical_path: &Path,
    mut output: W,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
    W: Write + Send,
{
    sort_handlers(handlers);
    let workbook_context = WriteWorkbookContext::new(logical_path);
    before_workbook(handlers, &workbook_context)?;

    if template_write::has_template(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    ) {
        write_xlsx_onto_template_package::<T, I>(
            logical_path,
            Some(&mut output),
            options,
            rows,
            handlers,
        )?;
    } else {
        let mut workbook = Workbook::new();
        write_sheet_to_workbook::<T, I>(&mut workbook, options, rows, handlers)?;
        save_workbook_to_writer(&mut workbook, &mut output, options.password.as_deref())?;
    }
    after_workbook(handlers, &workbook_context)
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
    W: Write + Send + 'static,
{
    validate_csv_options(options)?;
    write_csv_to::<T, I>(logical_path, Box::new(output), options, rows, handlers)
}

/// Builds a complete CSV document in memory.
///
/// This is primarily used when a borrowed output stream must not receive a
/// partial document if row conversion or a handler fails.
///
/// # Errors
///
/// Returns a conversion, handler, CSV-format, or charset error.
pub fn write_csv_to_buffer<T, I>(
    logical_path: &Path,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<Vec<u8>>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let output = CapturedOutput::default();
    write_csv_to_writer::<T, I, _>(logical_path, output.clone(), options, rows, handlers)?;
    take_captured_output(&output)
}

#[derive(Clone, Default)]
struct CapturedOutput(Arc<Mutex<Vec<u8>>>);

impl Write for CapturedOutput {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.0
            .lock()
            .map_err(|_| std::io::Error::other("CSV capture lock poisoned"))?
            .extend_from_slice(buffer);
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn take_captured_output(output: &CapturedOutput) -> Result<Vec<u8>> {
    let mut bytes = output
        .0
        .lock()
        .map_err(|_| ExcelError::Io(std::io::Error::other("CSV capture lock poisoned")))?;
    Ok(std::mem::take(&mut *bytes))
}

fn write_csv_to<T, I>(
    path: &Path,
    output: Box<dyn Write + Send>,
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
    output: Box<dyn Write + Send>,
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
    mut output: Box<dyn Write + Send>,
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

/// Incremental UTF-8 to configured CSV charset transcoder.
///
/// This is the low-level counterpart of Java's charset-aware CSV output path.
/// Call [`Self::finish`] after the last chunk so incomplete UTF-8 and encoder
/// finalization errors are reported.
pub struct CsvEncodingWriter {
    output: Box<dyn Write + Send>,
    encoder: CsvEncoder,
    pending_utf8: Vec<u8>,
}

impl CsvEncodingWriter {
    /// Creates a transcoding writer for a Java-style charset name.
    ///
    /// # Errors
    ///
    /// Returns an error when the charset is unsupported.
    pub fn with_charset<W>(output: W, charset: &CsvCharset) -> Result<Self>
    where
        W: Write + Send + 'static,
    {
        Ok(Self::new(Box::new(output), csv_encoding(charset)?))
    }

    fn new(output: Box<dyn Write + Send>, encoding: CsvEncoding) -> Self {
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

    /// Finalizes the charset encoder and flushes the underlying output.
    ///
    /// # Errors
    ///
    /// Returns an error for incomplete UTF-8 or an underlying output failure.
    pub fn finish(&mut self) -> std::io::Result<()> {
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

/// Saves a workbook to `path` (optionally password-protected).
///
/// `pub(crate)` so executor integration tests can persist worksheets built via
/// [`ExcelWriteAddExecutor`] without duplicating the save path.
pub(crate) fn save_workbook(
    workbook: &mut Workbook,
    path: &Path,
    password: Option<&str>,
) -> Result<()> {
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

fn save_workbook_to_writer(
    workbook: &mut Workbook,
    output: &mut (dyn Write + Send),
    password: Option<&str>,
) -> Result<()> {
    if let Some(password) = password {
        let mut encrypted = std::io::Cursor::new(Vec::new());
        save_encrypted_workbook_to(workbook, password, &mut encrypted)?;
        output.write_all(encrypted.get_ref())?;
    } else {
        workbook
            .save_to_writer(&mut *output)
            .map_err(format_error)?;
    }
    output.flush()?;
    Ok(())
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
            relative_row_index: Some(usize::try_from(row_index).unwrap_or(0)),
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
            relative_row_index: None,
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

/// Tracks the next physical row / data-row index while appending.
///
/// Used by [`ExcelWriteAddExecutor`] and the stateful [`ExcelWriter`] path that
/// both delegate to [`append_rows_to_worksheet`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WriteProgress {
    /// Next 0-based physical worksheet row to write.
    pub next_row: u32,
    /// Next 0-based data-row index (excludes header rows).
    pub next_data_index: usize,
}

#[derive(Clone, Copy)]
struct SheetStyleContext<'a> {
    explicit: Option<&'a CellStyle>,
    metadata: &'a ExcelWriteMetadata,
    is_head: bool,
    global: WriteGlobalFlags,
}

impl<'a> SheetStyleContext<'a> {
    const fn head(
        explicit: &'a CellStyle,
        metadata: &'a ExcelWriteMetadata,
        global: WriteGlobalFlags,
    ) -> Self {
        Self {
            explicit: Some(explicit),
            metadata,
            is_head: true,
            global,
        }
    }

    const fn content(
        explicit: Option<&'a CellStyle>,
        metadata: &'a ExcelWriteMetadata,
        global: WriteGlobalFlags,
    ) -> Self {
        Self {
            explicit,
            metadata,
            is_head: false,
            global,
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
            handler_cell: None,
            global: self.global,
        }
    }
}

#[derive(Clone, Copy)]
struct CellFormatContext<'a> {
    explicit: Option<&'a CellStyle>,
    cell: Option<ExcelCellStyle>,
    font: Option<ExcelFontStyle>,
    /// Style contributed by registered WriteHandler strategies
    /// (Java `AbstractCellStyleStrategy` merge into `WriteCellData`).
    handler_cell: Option<ExcelCellStyle>,
    global: WriteGlobalFlags,
}

impl CellFormatContext<'_> {
    /// Attaches a strategy-derived cell style (Java `WriteCellStyle.merge`).
    #[must_use]
    const fn with_handler_cell(mut self, handler_cell: Option<ExcelCellStyle>) -> Self {
        self.handler_cell = handler_cell;
        self
    }
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
    /// Builds image pixel layout from explicit options, annotation widths, and
    /// registered column-width strategies
    /// (Java `SimpleColumnWidthStyleStrategy` / `AbstractColumnWidthStyleStrategy`).
    ///
    /// Precedence: explicit `WriteOptions` widths win; registered handler
    /// strategies overwrite annotation/`@ColumnWidth` values for schema
    /// columns. Columns outside the typed schema keep Excel default `64` px.
    fn new(
        columns: &[(usize, usize, &'static ExcelColumn)],
        options: &WriteOptions,
        metadata: &ExcelWriteMetadata,
        head_rows: u32,
        handlers: &[Box<dyn WriteHandler>],
    ) -> Result<Self> {
        let mut column_widths = HashMap::new();
        // Explicit WriteOptions widths win (same precedence as sheet write path).
        for (column, width) in &options.column_widths {
            column_widths.insert(*column, excel_column_width_pixels(*width));
        }
        // Annotation `@ColumnWidth` / type-level column width.
        for (physical_index, _, column) in columns {
            let physical_index = to_column(*physical_index)?;
            if column_widths.contains_key(&physical_index) {
                continue;
            }
            if let Some(width) = column.column_width.or(metadata.column_width) {
                column_widths.insert(physical_index, excel_column_width_pixels(width));
            }
        }
        // Registered handler strategies override annotation widths so image
        // pixel layout matches `apply_handler_column_widths` (Java
        // `SimpleColumnWidthStyleStrategy` / `setColumnWidth` after annotations).
        for (physical_index, _, _) in columns {
            let physical_index = to_column(*physical_index)?;
            if options
                .column_widths
                .iter()
                .any(|(explicit, _)| *explicit == physical_index)
            {
                continue;
            }
            for handler in handlers {
                if let Some(width) = handler.style_column_width(usize::from(physical_index)) {
                    column_widths.insert(physical_index, excel_column_width_pixels(width));
                }
            }
        }
        Ok(Self {
            column_widths,
            head_rows,
            head_row_height: excel_row_height_pixels(
                metadata
                    .head_row_height
                    .or_else(|| collect_handler_head_row_height(handlers)),
            ),
            content_row_height: excel_row_height_pixels(
                metadata
                    .content_row_height
                    .or_else(|| collect_handler_content_row_height(handlers)),
            ),
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


/// Sets an OOXML column width that serializes as exact character units.
///
/// Java / POI `Sheet.setColumnWidth(col, chars * 256)` becomes
/// `width="{chars}"` in worksheet XML. `rust_xlsxwriter`'s
/// [`Worksheet::set_column_width`] stores `chars * 7 + 5` pixels and round-trips
/// to `~chars + 0.71`; using `chars * 7` pixels yields exact `width="{chars}"`.
fn set_xlsx_column_width_chars(
    worksheet: &mut Worksheet,
    column: u16,
    chars: u16,
) -> Result<()> {
    let pixels = u32::from(chars).saturating_mul(7);
    worksheet
        .set_column_width_pixels(column, pixels)
        .map_err(format_error)?;
    Ok(())
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
    let mut spill = if options.compress_temp_files {
        Some(gzip_spill::GzipSheetDataWriter::create_owned(&options.sheet_name)?)
    } else {
        None
    };
    write_sheet_to_workbook_with_gzip::<T, I>(
        workbook,
        options,
        rows,
        handlers,
        spill.as_mut(),
    )
}

/// Creates a worksheet and appends rows, optionally mirroring into a gzip spill.
fn write_sheet_to_workbook_with_gzip<T, I>(
    workbook: &mut Workbook,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
    gzip_spill: Option<&mut gzip_spill::GzipSheetDataWriter>,
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let worksheet = if uses_constant_memory_spill(options) {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(&options.sheet_name)
        .map_err(format_error)?;
    for (column, width) in &options.column_widths {
        set_xlsx_column_width_chars(worksheet, *column, *width)?;
    }
    apply_annotation_column_widths::<T>(worksheet, options)?;
    // Static strategy widths (e.g. SimpleColumnWidth) apply before cells.
    apply_handler_column_widths::<T>(worksheet, options, handlers)?;
    apply_annotation_once_absolute_merge::<T>(worksheet)?;
    // Java `OnceAbsoluteMergeStrategy.afterSheetCreate` via registerWriteHandler
    apply_handler_once_absolute_merge(worksheet, handlers, 0)?;
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

    let progress = append_rows_to_worksheet_with_gzip::<T, I>(
        worksheet,
        options,
        rows,
        handlers,
        WriteProgress {
            // Java `WriteContextImpl.initHead`: newRowIndex += relativeHeadRowIndex()
            next_row: relative_head_start_row(options),
            next_data_index: 0,
        },
        true,
        T::write_metadata(),
        gzip_spill,
    )?;
    after_sheet(handlers, &sheet_context)?;
    // Optional autofit first; byte-length widths reapplied so LongestMatch
    // is not autofit-only (Java setColumnWidth(String.getBytes().length)).
    if options.auto_width || handlers_request_auto_width(handlers) {
        worksheet.autofit();
    }
    // LongestMatch measures in after_cell — re-apply measured widths after write
    // (Java AbstractColumnWidthStyleStrategy.afterCellDispose → setColumnWidth).
    apply_handler_column_widths::<T>(worksheet, options, handlers)?;
    Ok(progress)
}

/// ZIP/OOXML `withTemplate` path: preserve styles/merges and append sheetData.
///
/// When the requested sheet is missing, creates a new worksheet part inside the
/// package so existing sheets keep their styles and merges. The legacy
/// calamine → `rust_xlsxwriter` seed path is used only when
/// [`WriteOptions::use_legacy_template_seed`] is set.
fn write_xlsx_onto_template_package<T, I>(
    path: &Path,
    output: Option<&mut (dyn Write + Send)>,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    template_write::validate_template_source(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    )?;
    let bytes = template_write::load_template_bytes(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    )?;
    if options.use_legacy_template_seed {
        let mut workbook = Workbook::new();
        write_sheet_onto_template::<T, I>(&mut workbook, options, rows, handlers)?;
        return match output {
            Some(writer) => {
                save_workbook_to_writer(&mut workbook, writer, options.password.as_deref())
            }
            None => save_workbook(&mut workbook, path, options.password.as_deref()),
        };
    }

    let mut package = template_write::TemplatePackage::from_bytes(&bytes)?;
    let sheet_names = package.sheet_names()?;
    let (_target_index, target_name, create_new) = template_write::resolve_package_target(
        &sheet_names,
        options.sheet_index,
        &options.sheet_name,
    );
    if create_new {
        package.ensure_sheet(&target_name)?;
    }

    let mut write_options = options.clone();
    write_options.sheet_name = target_name.clone();
    let append_rows = collect_template_append_rows::<T, I>(&write_options, rows, true, 0)?;
    let sheet_context = WriteSheetContext::new(&target_name);
    before_sheet(handlers, &sheet_context)?;
    package.append_rows(&target_name, &append_rows)?;
    after_sheet(handlers, &sheet_context)?;
    save_template_package(&package, path, output, options.password.as_deref())
}

/// Builds sparse `(physical_column, value)` rows for ZIP `sheetData` append.
fn collect_template_append_rows<T, I>(
    options: &WriteOptions,
    rows: I,
    write_head: bool,
    _next_data_index: usize,
) -> Result<Vec<Vec<(usize, CellValue)>>>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let columns = selected_columns(T::schema(), options);
    let mut output = Vec::new();
    let head_rows = head_rows_for_schema(T::schema(), options)?;
    if write_head && head_rows > 0 {
        if let Some(head) = &options.dynamic_head {
            for level in 0..usize::try_from(head_rows).unwrap_or(0) {
                let mut row = Vec::with_capacity(columns.len());
                for (physical_index, _, _) in &columns {
                    let label = head
                        .get(*physical_index)
                        .and_then(|path| path.get(level))
                        .cloned()
                        .unwrap_or_default();
                    row.push((*physical_index, CellValue::String(label)));
                }
                output.push(row);
            }
        } else {
            let mut row = Vec::with_capacity(columns.len());
            for (physical_index, _, column) in &columns {
                row.push((*physical_index, CellValue::String(column.name.to_owned())));
            }
            output.push(row);
        }
    }
    for row in rows {
        let cells = row.to_row_with_converters(&options.converters)?;
        let dynamic_columns = dynamic_columns_for_row(T::schema().is_empty(), cells.len(), options);
        let row_columns = dynamic_columns.as_deref().unwrap_or(&columns);
        let mut sparse = Vec::with_capacity(row_columns.len());
        for (physical_index, schema_index, _) in row_columns {
            let value = cells
                .get(*schema_index)
                .cloned()
                .unwrap_or(CellValue::Empty);
            sparse.push((*physical_index, value));
        }
        output.push(sparse);
    }
    Ok(output)
}

/// Persists a template package to a path or stream, optionally encrypting.
fn save_template_package(
    package: &template_write::TemplatePackage,
    path: &Path,
    output: Option<&mut (dyn Write + Send)>,
    password: Option<&str>,
) -> Result<()> {
    let plaintext = package.to_bytes()?;
    if let Some(password) = password {
        let mut encrypted = std::io::Cursor::new(Vec::new());
        save_encrypted_bytes_to(&plaintext, password, &mut encrypted)?;
        if let Some(writer) = output {
            writer.write_all(encrypted.get_ref())?;
            writer.flush()?;
        } else {
            std::fs::write(path, encrypted.get_ref())?;
        }
        return Ok(());
    }
    if let Some(writer) = output {
        writer.write_all(&plaintext)?;
        writer.flush()?;
        Ok(())
    } else {
        std::fs::write(path, plaintext).map_err(ExcelError::from)
    }
}

fn save_encrypted_bytes_to(
    plaintext: &[u8],
    password: &str,
    file: &mut dyn ReadWriteSeek,
) -> Result<()> {
    let mut random = rand::rng();
    Ecma376AgileWriter::create(&mut random, password, file)
        .map_err(ExcelError::from)
        .and_then(|mut writer| {
            let _ = writer.write_all(plaintext);
            writer.finalize().map_err(ExcelError::from)
        })
}

/// Seeds a workbook from `withTemplate` then appends typed rows to the target sheet.
///
/// **Legacy path only** — enabled via [`WriteOptions::use_legacy_template_seed`].
/// Value replay does not preserve styles/merges; prefer the ZIP package path.
///
/// Mirrors Java `WorkBookUtil.createWorkBook` (template branch) + `ExcelWriteAddExecutor`.
///
/// # Errors
///
/// Returns template validation/load errors, or standard XLSX write errors.
fn write_sheet_onto_template<T, I>(
    workbook: &mut Workbook,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    template_write::validate_template_source(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    )?;
    let bytes = template_write::load_template_bytes(
        options.template_file.as_deref(),
        options.template_bytes.as_deref(),
    )?;
    let sheets = template_write::load_template_sheets(&bytes)?;
    let (target_index, target_name, create_new) = template_write::resolve_template_target(
        &sheets,
        options.sheet_index,
        &options.sheet_name,
    );
    template_write::seed_workbook_from_template(workbook, &sheets)?;

    let mut write_options = options.clone();
    write_options.sheet_name = target_name.clone();

    if create_new {
        // Java creates a new sheet when the requested name/index is absent.
        return write_sheet_to_workbook::<T, I>(workbook, &write_options, rows, handlers);
    }

    let start_row = sheets
        .get(target_index)
        .map(|sheet| sheet.next_row)
        .unwrap_or(0);
    let worksheet = workbook
        .worksheet_from_name(&target_name)
        .map_err(format_error)?;
    for (column, width) in &write_options.column_widths {
        set_xlsx_column_width_chars(worksheet, *column, *width)?;
    }
    apply_annotation_column_widths::<T>(worksheet, &write_options)?;
    apply_handler_column_widths::<T>(worksheet, &write_options, handlers)?;
    apply_annotation_once_absolute_merge_at::<T>(worksheet, start_row)?;
    apply_handler_once_absolute_merge(worksheet, handlers, start_row)?;
    for range in &write_options.merge_ranges {
        let offset = start_row;
        worksheet
            .merge_range(
                range.first_row.saturating_add(offset),
                range.first_column,
                range.last_row.saturating_add(offset),
                range.last_column,
                "",
                &Format::new(),
            )
            .map_err(format_error)?;
    }

    let sheet_context = WriteSheetContext::new(&target_name);
    before_sheet(handlers, &sheet_context)?;
    let mut spill = if write_options.compress_temp_files {
        Some(gzip_spill::GzipSheetDataWriter::create_owned(&target_name)?)
    } else {
        None
    };
    let progress = append_rows_to_worksheet_with_gzip::<T, I>(
        worksheet,
        &write_options,
        rows,
        handlers,
        WriteProgress {
            next_row: start_row,
            next_data_index: 0,
        },
        true,
        T::write_metadata(),
        spill.as_mut(),
    )?;
    after_sheet(handlers, &sheet_context)?;
    if write_options.auto_width || handlers_request_auto_width(handlers) {
        worksheet.autofit();
    }
    // Byte-length widths win over optional autofit fallback.
    apply_handler_column_widths::<T>(worksheet, &write_options, handlers)?;
    Ok(progress)
}

/// Appends typed rows onto an existing worksheet.
///
/// Java counterpart: the body of `ExcelWriteAddExecutor.add(Collection<?>)`
/// plus `addOneRowOfDataToExcel` (header / cell / handler orchestration).
/// Kept here so the historical `lib.rs` writer path stays intact; the
/// mirrored executor delegates to this function (只增不减).
///
/// # Errors
///
/// Returns a conversion, handler, or XLSX-format error.
pub fn append_rows_to_worksheet<T, I>(
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
    append_rows_to_worksheet_with_gzip::<T, I>(
        worksheet,
        options,
        rows,
        handlers,
        progress,
        write_head,
        metadata,
        None,
    )
}

/// Like [`append_rows_to_worksheet`], optionally mirroring data rows into a gzip spill.
///
/// Java mapping: when `compress_temp_files` is on, [`gzip_spill::GzipSheetDataWriter`]
/// mirrors POI `GZIPSheetDataWriter` for observability and disk spill.
#[allow(clippy::too_many_arguments)]
pub fn append_rows_to_worksheet_with_gzip<T, I>(
    worksheet: &mut Worksheet,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
    progress: WriteProgress,
    write_head: bool,
    metadata: &ExcelWriteMetadata,
    mut gzip_spill: Option<&mut gzip_spill::GzipSheetDataWriter>,
) -> Result<WriteProgress>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let WriteProgress {
        next_row: mut row_index,
        next_data_index: mut data_index,
    } = progress;
    let global = WriteGlobalFlags::from(options);
    let columns = selected_columns(T::schema(), options);
    let annotation_loop_merges = annotation_loop_merges_from_columns(&columns)?;
    let head_rows = head_rows_for_schema(T::schema(), options)?;
    let image_layout = ImageLayout::new(&columns, options, metadata, head_rows, handlers)?;
    if write_head && head_rows > 0 {
        if let Some(head) = &options.dynamic_head {
            write_dynamic_headers_with_handlers(
                worksheet,
                &columns,
                head,
                &options.sheet_name,
                SheetStyleContext::head(&options.head_style, metadata, global),
                handlers,
                &image_layout,
                row_index,
            )?;
        } else {
            write_headers_with_handlers(
                worksheet,
                &columns,
                &options.sheet_name,
                SheetStyleContext::head(&options.head_style, metadata, global),
                handlers,
                &image_layout,
                row_index,
            )?;
        }
        // Annotation `@HeadRowHeight` or registered `SimpleRowHeightStyleStrategy`
        let head_height = metadata
            .head_row_height
            .or_else(|| collect_handler_head_row_height(handlers));
        if let Some(height) = head_height {
            for head_row in row_index..row_index + head_rows {
                worksheet
                    .set_row_height(head_row, height)
                    .map_err(format_error)?;
            }
        }
        row_index += head_rows;
    }
    for row in rows {
        // Annotation `@ContentRowHeight` or registered `SimpleRowHeightStyleStrategy`
        let content_height = metadata
            .content_row_height
            .or_else(|| collect_handler_content_row_height(handlers));
        if let Some(height) = content_height {
            worksheet
                .set_row_height(row_index, height)
                .map_err(format_error)?;
        }
        let cells = row.to_row_with_converters(&options.converters)?;
        if let Some(spill) = gzip_spill.as_mut() {
            spill.write_row(&cells)?;
        }
        let dynamic_columns = dynamic_columns_for_row(T::schema().is_empty(), cells.len(), options);
        let row_columns = dynamic_columns.as_deref().unwrap_or(&columns);
        let style = (!options.content_styles.is_empty())
            .then(|| &options.content_styles[data_index % options.content_styles.len()]);
        apply_loop_merges(worksheet, row_index, data_index, &options.loop_merges)?;
        apply_loop_merges(worksheet, row_index, data_index, &annotation_loop_merges)?;
        write_data_row_with_handlers(
            worksheet,
            row_index,
            data_index,
            row_columns,
            &cells,
            &options.sheet_name,
            SheetStyleContext::content(style, metadata, global),
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
            set_xlsx_column_width_chars(worksheet, to_column(physical_index)?, width)?;
        }
    }
    Ok(())
}

/// Applies column widths from registered strategies
/// (Java `SimpleColumnWidthStyleStrategy` / `AbstractColumnWidthStyleStrategy`).
fn apply_handler_column_widths<T>(
    worksheet: &mut Worksheet,
    options: &WriteOptions,
    handlers: &[Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
{
    for (physical_index, _, _) in selected_columns(T::schema(), options) {
        let column = to_column(physical_index)?;
        // Explicit `WriteOptions::column_widths` wins over strategies.
        if options
            .column_widths
            .iter()
            .any(|(explicit, _)| *explicit == column)
        {
            continue;
        }
        for handler in handlers {
            if let Some(width) = handler.style_column_width(physical_index) {
                set_xlsx_column_width_chars(worksheet, column, width)?;
            }
        }
    }
    Ok(())
}

/// Collects head row height from registered strategies
/// (Java `SimpleRowHeightStyleStrategy`).
fn collect_handler_head_row_height(handlers: &[Box<dyn WriteHandler>]) -> Option<u16> {
    handlers
        .iter()
        .rev()
        .find_map(|handler| handler.style_head_row_height())
}

/// Collects content row height from registered strategies
/// (Java `SimpleRowHeightStyleStrategy`).
fn collect_handler_content_row_height(handlers: &[Box<dyn WriteHandler>]) -> Option<u16> {
    handlers
        .iter()
        .rev()
        .find_map(|handler| handler.style_content_row_height())
}

/// Whether any handler requests longest-match autofit
/// (Java `LongestMatchColumnWidthStyleStrategy`).
fn handlers_request_auto_width(handlers: &[Box<dyn WriteHandler>]) -> bool {
    handlers
        .iter()
        .any(|handler| handler.style_auto_column_width())
}

/// Merges cell styles from registered style strategies in handler order
/// (Java `AbstractCellStyleStrategy.afterCellDispose` + `WriteCellStyle.merge`).
fn collect_handler_cell_style(
    handlers: &[Box<dyn WriteHandler>],
    context: &WriteCellContext,
) -> Option<ExcelCellStyle> {
    let mut merged: Option<ExcelCellStyle> = None;
    for handler in handlers {
        if let Some(style) = handler.style_cell_style(context) {
            merged = Some(match merged {
                Some(target) => merge_write_cell_style(&style, target),
                None => style,
            });
        }
    }
    merged
}

/// Applies type-level `@OnceAbsoluteMerge` metadata when all indexes are non-negative.
fn apply_annotation_once_absolute_merge<T>(worksheet: &mut Worksheet) -> Result<()>
where
    T: ExcelRow,
{
    apply_annotation_once_absolute_merge_at::<T>(worksheet, 0)
}

/// Applies `@OnceAbsoluteMerge` with a row offset (template append writes).
fn apply_annotation_once_absolute_merge_at<T>(
    worksheet: &mut Worksheet,
    row_offset: u32,
) -> Result<()>
where
    T: ExcelRow,
{
    let Some(merge) = T::write_metadata().once_absolute_merge else {
        return Ok(());
    };
    apply_once_absolute_merge_property(worksheet, merge, row_offset)
}

/// Applies registered [`OnceAbsoluteMergeStrategy`] regions
/// (Java `OnceAbsoluteMergeStrategy.afterSheetCreate` → `addMergedRegionUnsafe`).
fn apply_handler_once_absolute_merge(
    worksheet: &mut Worksheet,
    handlers: &[Box<dyn WriteHandler>],
    row_offset: u32,
) -> Result<()> {
    for handler in handlers {
        if let Some(merge) = handler.style_once_absolute_merge() {
            apply_once_absolute_merge_property(worksheet, merge, row_offset)?;
        }
    }
    Ok(())
}

/// Shared absolute-merge apply used by annotation and registered strategy paths.
fn apply_once_absolute_merge_property(
    worksheet: &mut Worksheet,
    merge: easyexcel_core::OnceAbsoluteMergeProperty,
    row_offset: u32,
) -> Result<()> {
    if merge.first_row_index < 0
        || merge.last_row_index < 0
        || merge.first_column_index < 0
        || merge.last_column_index < 0
    {
        return Ok(());
    }
    #[allow(clippy::cast_sign_loss, clippy::cast_possible_truncation)]
    worksheet
        .merge_range(
            (merge.first_row_index as u32).saturating_add(row_offset),
            merge.first_column_index as u16,
            (merge.last_row_index as u32).saturating_add(row_offset),
            merge.last_column_index as u16,
            "",
            &Format::new(),
        )
        .map_err(format_error)?;
    Ok(())
}

/// Builds loop-merge strategies from field-level `@ContentLoopMerge` metadata.
fn annotation_loop_merges_from_columns(
    columns: &[(usize, usize, &'static ExcelColumn)],
) -> Result<Vec<LoopMergeStrategy>> {
    let mut strategies = Vec::new();
    for (physical_index, _, column) in columns {
        let Some(property) = column.loop_merge else {
            continue;
        };
        strategies.push(LoopMergeStrategy::new(
            property.each_row,
            property.column_extend,
            to_column(*physical_index)?,
        )?);
    }
    Ok(strategies)
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

/// Java `relativeHeadRowIndex` → zero-based start row for a new sheet write.
fn relative_head_start_row(options: &WriteOptions) -> u32 {
    if options.relative_head_row_index <= 0 {
        0
    } else {
        u32::try_from(options.relative_head_row_index).unwrap_or(0)
    }
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
        SheetStyleContext::head(&CellStyle::new(), &METADATA, WriteGlobalFlags::default()),
        &mut [],
        &layout,
        0,
    )
}

fn write_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    sheet_name: &str,
    style: SheetStyleContext<'_>,
    handlers: &mut [Box<dyn WriteHandler>],
    image_layout: &ImageLayout,
    start_row: u32,
) -> Result<()> {
    let labels = columns
        .iter()
        .map(|(_, _, column)| column.name.to_owned())
        .collect::<Vec<_>>();
    write_header_row_with_handlers(
        worksheet,
        start_row,
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
    start_row: u32,
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
        let row_index = start_row.saturating_add(level as u32);
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
        let column_index = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index,
            field: (!column.field.is_empty()).then_some(column.field),
            is_head: true,
            relative_row_index: Some(usize::try_from(row_index).unwrap_or(0)),
            value: CellValue::String(label.clone()),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            let handler_style = collect_handler_cell_style(handlers, &context);
            let format_context = style
                .column(column)
                .with_handler_cell(handler_style);
            let format = cell_format(format_context);
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
        0,
        columns,
        cells,
        "",
        SheetStyleContext {
            explicit: None,
            metadata: &ExcelWriteMetadata::new(),
            is_head: false,
            global: WriteGlobalFlags::default(),
        },
        &mut [],
        &image_layout,
    )
}

#[allow(clippy::too_many_arguments)]
fn write_data_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    relative_row_index: usize,
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
            relative_row_index: Some(relative_row_index),
            value: value.clone(),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            let handler_style = collect_handler_cell_style(handlers, &context);
            let format_context = style
                .column(metadata)
                .with_handler_cell(handler_style);
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

#[allow(clippy::too_many_lines)]
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
    let global = style.global;
    match value {
        CellValue::Empty => {
            worksheet
                .write_blank(row_index, column, &format)
                .map_err(format_error)?;
        }
        CellValue::String(value) | CellValue::Error(value) => {
            let text = maybe_trim_cell_string(value, global.auto_trim);
            worksheet
                .write_string_with_format(row_index, column, &text, &format)
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
            let mut cell_format = format.clone();
            if global.use_scientific_format
                && metadata.format.is_none()
                && is_scientific_magnitude(*value)
            {
                cell_format = cell_format.set_num_format("0.#####E0");
            }
            worksheet
                .write_number_with_format(row_index, column, *value, &cell_format)
                .map_err(format_error)?;
        }
        CellValue::Decimal(value) => {
            let value = value
                .to_f64()
                .filter(|value| value.is_finite())
                .ok_or_else(|| {
                    ExcelError::Format("decimal value exceeds XLSX numeric range".to_owned())
                })?;
            let mut cell_format = format.clone();
            if global.use_scientific_format
                && metadata.format.is_none()
                && is_scientific_magnitude(value)
            {
                cell_format = cell_format.set_num_format("0.#####E0");
            }
            worksheet
                .write_number_with_format(row_index, column, value, &cell_format)
                .map_err(format_error)?;
        }
        CellValue::Date(value) => {
            let format = format
                .clone()
                .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd"));
            if global.use_1904_windowing {
                let serial = date_to_excel_serial_with_windowing(*value, true);
                worksheet
                    .write_number_with_format(row_index, column, serial, &format)
                    .map_err(format_error)?;
            } else {
                worksheet
                    .write_datetime_with_format(row_index, column, *value, &format)
                    .map_err(format_error)?;
            }
        }
        CellValue::DateTime(value) => {
            let format = format
                .clone()
                .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd hh:mm:ss"));
            if global.use_1904_windowing {
                let serial = datetime_to_excel_serial_with_windowing(*value, true);
                worksheet
                    .write_number_with_format(row_index, column, serial, &format)
                    .map_err(format_error)?;
            } else {
                worksheet
                    .write_datetime_with_format(row_index, column, *value, &format)
                    .map_err(format_error)?;
            }
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
        CellValue::RichText(value) => {
            write_rich_text(worksheet, row_index, column, value, &format)?;
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

fn write_rich_text(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    data: &RichTextStringData,
    cell_format: &Format,
) -> Result<()> {
    if data.text_string().is_empty() {
        worksheet
            .write_string_with_format(row, column, "", cell_format)
            .map(|_| ())
            .map_err(format_error)?;
        return Ok(());
    }
    let runs = rich_text_runs(data)?;
    let references = runs
        .iter()
        .map(|(format, text)| (format, text.as_str()))
        .collect::<Vec<_>>();
    worksheet
        .write_rich_string_with_format(row, column, &references, cell_format)
        .map(|_| ())
        .map_err(format_error)
}

fn rich_text_runs(data: &RichTextStringData) -> Result<Vec<(Format, String)>> {
    let text = data.text_string();
    let utf16_length = text.encode_utf16().count();
    let mut boundaries = vec![0, utf16_length];
    for interval in data.interval_fonts() {
        let start = interval.start_index();
        let end = interval.end_index();
        if start >= end || end > utf16_length {
            return Err(ExcelError::Format(format!(
                "rich-text font range [{start}, {end}) is outside UTF-16 length {utf16_length}"
            )));
        }
        if utf16_byte_index(text, start).is_none() || utf16_byte_index(text, end).is_none() {
            return Err(ExcelError::Format(format!(
                "rich-text font range [{start}, {end}) splits a UTF-16 surrogate pair"
            )));
        }
        boundaries.push(start);
        boundaries.push(end);
    }
    boundaries.sort_unstable();
    boundaries.dedup();

    boundaries
        .windows(2)
        .map(|window| {
            let start = window[0];
            let end = window[1];
            let start_byte = utf16_byte_index(text, start).expect("validated UTF-16 boundary");
            let end_byte = utf16_byte_index(text, end).expect("validated UTF-16 boundary");
            let font = data
                .interval_fonts()
                .iter()
                .rev()
                .find(|interval| interval.start_index() <= start && interval.end_index() >= end)
                .map_or(data.write_font(), |interval| Some(interval.write_font()));
            Ok((
                font.map_or_else(Format::new, rich_text_format),
                text[start_byte..end_byte].to_owned(),
            ))
        })
        .collect()
}

fn utf16_byte_index(text: &str, target: usize) -> Option<usize> {
    let mut utf16_index = 0;
    for (byte_index, character) in text.char_indices() {
        if utf16_index == target {
            return Some(byte_index);
        }
        utf16_index += character.len_utf16();
        if utf16_index > target {
            return None;
        }
    }
    (utf16_index == target).then_some(text.len())
}

fn rich_text_format(font: &WriteFont) -> Format {
    let mut format = Format::new();
    if let Some(name) = font.get_font_name() {
        format = format.set_font_name(name);
    }
    if let Some(size) = font.get_font_height_in_points() {
        format = format.set_font_size(size);
    }
    if let Some(italic) = font.get_italic() {
        format = if italic {
            format.set_italic()
        } else {
            format.unset_italic()
        };
    }
    if let Some(strikeout) = font.get_strikeout() {
        format = if strikeout {
            format.set_font_strikethrough()
        } else {
            format.unset_font_strikethrough()
        };
    }
    if let Some(color) = font.get_color() {
        format = format.set_font_color(annotation_color(color));
    }
    if let Some(script) = font.get_type_offset() {
        format = format.set_font_script(annotation_font_script(script));
    }
    if let Some(underline) = font.get_underline() {
        format = format.set_underline(annotation_underline(underline));
    }
    if let Some(charset) = font.get_charset() {
        format = format.set_font_charset(charset);
    }
    if let Some(bold) = font.get_bold() {
        format = if bold {
            format.set_bold()
        } else {
            format.unset_bold()
        };
    }
    format
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
    // Annotation style merged with handler strategy style
    // (Java `WriteCellStyle.merge(strategy, cellData.getOrCreateStyle())`).
    let mut annotation_cell = context.cell;
    if let Some(handler_style) = context.handler_cell {
        annotation_cell = Some(merge_write_cell_style(
            &handler_style,
            annotation_cell.unwrap_or_default(),
        ));
    }
    // Nested WriteFont / ExcelFontStyle on merged cell style
    // (Java WriteCellStyle.writeFont merge onto annotation HeadFontStyle/ContentFontStyle).
    let mut font = context.font;
    if let Some(style) = annotation_cell {
        if let Some(style_font) = style.font {
            font = Some(match font {
                Some(target) => merge_handler_font_style(&style_font, target),
                None => style_font,
            });
        }
        format = apply_annotation_cell_style(format, style);
    }
    if let Some(font) = font {
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
    // Nested WriteFont / ExcelFontStyle (Java WriteCellStyle.writeFont)
    if let Some(font) = style.font {
        format = apply_annotation_font_style(format, font);
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

pub(crate) fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[cfg(test)]
mod tests;
#[cfg(test)]
mod missing_tests;
