//! OOXML-preserving XLSX template filling.

mod builder_fill_executor;

pub use builder_fill_executor::{BuilderFillExecutor, create_builder_fill_executor};

use std::any::Any;
use std::collections::BTreeMap;
use std::fmt::Write as FmtWrite;
use std::fs::File;
use std::io::{Cursor, Read, Seek, Write};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::{Path, PathBuf};

use bigdecimal::BigDecimal;
use chrono::{NaiveDate, NaiveDateTime};
use easyexcel_core::{CellValue, ExcelError, Result};
use easyexcel_writer::ExcelOutputStream;
use num_bigint::BigInt;
use zip::CompressionMethod;
use zip::read::ZipArchive;
use zip::write::{SimpleFileOptions, ZipWriter};

/// Value accepted by [`TemplateData`] placeholder insertion methods.
pub trait IntoTemplateValue {
    /// Converts the value to its typed template representation.
    fn into_template_value(self) -> CellValue;
}

impl IntoTemplateValue for CellValue {
    fn into_template_value(self) -> CellValue {
        self
    }
}

impl IntoTemplateValue for String {
    fn into_template_value(self) -> CellValue {
        CellValue::String(self)
    }
}

impl IntoTemplateValue for &str {
    fn into_template_value(self) -> CellValue {
        CellValue::String(self.to_owned())
    }
}

impl IntoTemplateValue for &String {
    fn into_template_value(self) -> CellValue {
        CellValue::String(self.clone())
    }
}

impl IntoTemplateValue for bool {
    fn into_template_value(self) -> CellValue {
        CellValue::Bool(self)
    }
}

macro_rules! impl_integer_template_value {
    ($($type:ty),+ $(,)?) => {
        $(
            impl IntoTemplateValue for $type {
                fn into_template_value(self) -> CellValue {
                    CellValue::Int(i64::from(self))
                }
            }
        )+
    };
}

impl_integer_template_value!(i8, i16, i32, i64, u8, u16, u32);

macro_rules! impl_decimal_integer_template_value {
    ($($type:ty),+ $(,)?) => {
        $(
            impl IntoTemplateValue for $type {
                fn into_template_value(self) -> CellValue {
                    CellValue::Decimal(BigDecimal::from(self))
                }
            }
        )+
    };
}

impl_decimal_integer_template_value!(i128, u64, u128);

impl IntoTemplateValue for isize {
    fn into_template_value(self) -> CellValue {
        CellValue::Int(i64::try_from(self).expect("Rust isize is at most 64 bits"))
    }
}

impl IntoTemplateValue for usize {
    fn into_template_value(self) -> CellValue {
        CellValue::Decimal(BigDecimal::from(
            u64::try_from(self).expect("Rust usize is at most 64 bits"),
        ))
    }
}

impl IntoTemplateValue for BigInt {
    fn into_template_value(self) -> CellValue {
        CellValue::Decimal(BigDecimal::from(self))
    }
}

impl IntoTemplateValue for f32 {
    fn into_template_value(self) -> CellValue {
        CellValue::Float(f64::from(self))
    }
}

impl IntoTemplateValue for f64 {
    fn into_template_value(self) -> CellValue {
        CellValue::Float(self)
    }
}

impl IntoTemplateValue for BigDecimal {
    fn into_template_value(self) -> CellValue {
        CellValue::Decimal(self)
    }
}

impl IntoTemplateValue for NaiveDate {
    fn into_template_value(self) -> CellValue {
        CellValue::Date(self)
    }
}

impl IntoTemplateValue for NaiveDateTime {
    fn into_template_value(self) -> CellValue {
        CellValue::DateTime(self)
    }
}

impl<T> IntoTemplateValue for Option<T>
where
    T: IntoTemplateValue,
{
    fn into_template_value(self) -> CellValue {
        self.map_or(CellValue::Empty, IntoTemplateValue::into_template_value)
    }
}

/// Scalar values used to replace `{key}` placeholders in OOXML text nodes.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TemplateData {
    values: BTreeMap<String, CellValue>,
}

/// Direction used when expanding a collection placeholder.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FillDirection {
    /// Repeats the template row downward.
    #[default]
    Vertical,
    /// Repeats the template cell to the right.
    Horizontal,
}

/// Collection fill behavior corresponding to Java `EasyExcel`'s `FillConfig`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FillConfig {
    direction: FillDirection,
    force_new_row: bool,
    auto_style: bool,
}

impl Default for FillConfig {
    fn default() -> Self {
        Self {
            direction: FillDirection::Vertical,
            force_new_row: false,
            auto_style: true,
        }
    }
}

impl FillConfig {
    /// Creates Java-compatible default fill configuration.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            direction: FillDirection::Vertical,
            force_new_row: false,
            auto_style: true,
        }
    }

    /// Sets vertical or horizontal collection expansion.
    #[must_use]
    pub const fn direction(mut self, direction: FillDirection) -> Self {
        self.direction = direction;
        self
    }

    /// Controls whether rows below a vertical template row are shifted.
    #[must_use]
    pub const fn force_new_row(mut self, force_new_row: bool) -> Self {
        self.force_new_row = force_new_row;
        self
    }

    /// Controls whether cloned cells retain the template cell style.
    #[must_use]
    pub const fn auto_style(mut self, auto_style: bool) -> Self {
        self.auto_style = auto_style;
        self
    }

    /// Returns the configured expansion direction.
    #[must_use]
    pub const fn get_direction(self) -> FillDirection {
        self.direction
    }

    /// Returns whether vertical filling shifts following rows.
    #[must_use]
    pub const fn get_force_new_row(self) -> bool {
        self.force_new_row
    }

    /// Returns whether template style is inherited.
    #[must_use]
    pub const fn get_auto_style(self) -> bool {
        self.auto_style
    }
}

/// Named or unnamed collection data corresponding to Java `EasyExcel`'s `FillWrapper`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FillWrapper {
    name: Option<String>,
    rows: Vec<TemplateData>,
}

impl FillWrapper {
    /// Creates an unnamed collection for `{.field}` placeholders.
    #[must_use]
    pub fn new(rows: impl IntoIterator<Item = TemplateData>) -> Self {
        Self {
            name: None,
            rows: rows.into_iter().collect(),
        }
    }

    /// Creates a named collection for `{name.field}` placeholders.
    #[must_use]
    pub fn named(name: impl Into<String>, rows: impl IntoIterator<Item = TemplateData>) -> Self {
        Self {
            name: Some(name.into()),
            rows: rows.into_iter().collect(),
        }
    }

    /// Returns the optional collection prefix.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns collection rows in fill order.
    #[must_use]
    pub fn rows(&self) -> &[TemplateData] {
        &self.rows
    }
}

impl TemplateData {
    /// Creates empty template data.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            values: BTreeMap::new(),
        }
    }

    /// Adds or replaces a placeholder value.
    #[must_use]
    pub fn with(mut self, key: impl Into<String>, value: impl IntoTemplateValue) -> Self {
        self.values.insert(key.into(), value.into_template_value());
        self
    }

    /// Inserts a placeholder value and returns the previous value.
    pub fn insert(
        &mut self,
        key: impl Into<String>,
        value: impl IntoTemplateValue,
    ) -> Option<CellValue> {
        self.values.insert(key.into(), value.into_template_value())
    }

    /// Returns all values in deterministic key order.
    #[must_use]
    pub const fn values(&self) -> &BTreeMap<String, CellValue> {
        &self.values
    }
}

#[derive(Debug)]
struct TemplateEntry {
    name: String,
    is_dir: bool,
    compression: CompressionMethod,
    unix_mode: Option<u32>,
    bytes: Vec<u8>,
}

#[derive(Debug, Clone)]
struct PendingCollectionFill {
    wrapper: FillWrapper,
    config: FillConfig,
}

/// Worksheet selected for Java-style template fill and write operations.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub enum TemplateSheet {
    /// Selects a worksheet by its zero-based workbook order.
    #[default]
    First,
    /// Selects a worksheet by its zero-based workbook order.
    Index(usize),
    /// Selects a worksheet by its exact workbook name.
    Name(String),
}

impl TemplateSheet {
    /// Selects the first worksheet, equivalent to Java `writerSheet().build()`.
    #[must_use]
    pub const fn first() -> Self {
        Self::First
    }

    /// Selects a worksheet by Java-style zero-based sheet number.
    #[must_use]
    pub const fn index(index: usize) -> Self {
        Self::Index(index)
    }

    /// Selects a worksheet by exact name.
    #[must_use]
    pub fn name(name: impl Into<String>) -> Self {
        Self::Name(name.into())
    }
}

#[derive(Debug)]
struct PendingSheetFill {
    sheet: TemplateSheet,
    scalar: TemplateData,
    collections: Vec<PendingCollectionFill>,
    appended_rows: Vec<Vec<CellValue>>,
}

#[derive(Debug)]
struct ResolvedSheetFill {
    worksheet: String,
    scalar: TemplateData,
    collections: Vec<PendingCollectionFill>,
    appended_rows: Vec<Vec<CellValue>>,
}

impl PendingSheetFill {
    fn new(sheet: TemplateSheet) -> Self {
        Self {
            sheet,
            scalar: TemplateData::new(),
            collections: Vec::new(),
            appended_rows: Vec::new(),
        }
    }
}

/// Stateful OOXML template writer matching Java `ExcelWriter.fill` lifecycle.
///
/// Scalar values and collection fills are accumulated against one loaded XLSX
/// package. Repeated collection fills with the same prefix append at the prior
/// fill position instead of reopening the original template.
pub struct ExcelTemplateWriter<'a> {
    output: TemplateOutput<'a>,
    entries: Vec<TemplateEntry>,
    sheets: Vec<PendingSheetFill>,
    finished: bool,
    auto_close_stream: bool,
}

enum TemplateOutput<'a> {
    Path(PathBuf),
    Borrowed(&'a mut dyn Write),
    Owned(Box<dyn CloseableWrite + 'a>),
}

trait CloseableWrite: Write {
    fn close(&self) -> std::io::Result<()>;
}

impl<W> CloseableWrite for ExcelOutputStream<W>
where
    W: Write,
{
    fn close(&self) -> std::io::Result<()> {
        ExcelOutputStream::close(self)
    }
}

impl std::fmt::Debug for ExcelTemplateWriter<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let output = match self.output {
            TemplateOutput::Path(_) => "path",
            TemplateOutput::Borrowed(_) => "borrowed stream",
            TemplateOutput::Owned(_) => "owned stream",
        };
        formatter
            .debug_struct("ExcelTemplateWriter")
            .field("output", &output)
            .field("entries", &self.entries)
            .field("sheets", &self.sheets)
            .field("finished", &self.finished)
            .field("auto_close_stream", &self.auto_close_stream)
            .finish()
    }
}

impl ExcelTemplateWriter<'static> {
    /// Loads a template package for stateful filling.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn new(template: impl AsRef<Path>, output: impl Into<PathBuf>) -> Result<Self> {
        Ok(Self::from_entries(
            TemplateOutput::Path(output.into()),
            load_entries(template.as_ref())?,
        ))
    }

    /// Loads a template from a Java-style input stream and writes to a path.
    ///
    /// The reader is consumed and dropped after its bytes have been copied into
    /// memory, matching Java `EasyExcel`'s default `autoCloseStream(true)` input
    /// lifecycle. Pass `&mut reader` to retain caller ownership.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn from_reader<R>(template: R, output: impl Into<PathBuf>) -> Result<Self>
    where
        R: Read,
    {
        Ok(Self::from_entries(
            TemplateOutput::Path(output.into()),
            load_entries_from_reader(Box::new(template))?,
        ))
    }
}

impl<'a> ExcelTemplateWriter<'a> {
    fn from_entries(output: TemplateOutput<'a>, entries: Vec<TemplateEntry>) -> Self {
        Self {
            output,
            entries,
            sheets: vec![PendingSheetFill::new(TemplateSheet::first())],
            finished: false,
            auto_close_stream: true,
        }
    }

    /// Loads a path template and writes to a caller-owned output stream.
    ///
    /// The borrowed writer is flushed but never closed or dropped by this
    /// object, which is Rust's equivalent of Java `autoCloseStream(false)`.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn to_writer<W>(template: impl AsRef<Path>, output: &'a mut W) -> Result<Self>
    where
        W: Write,
    {
        Ok(Self::from_entries(
            TemplateOutput::Borrowed(output),
            load_entries(template.as_ref())?,
        ))
    }

    /// Loads a stream template and writes to a caller-owned output stream.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn from_reader_to_writer<R, W>(template: R, output: &'a mut W) -> Result<Self>
    where
        R: Read,
        W: Write,
    {
        Ok(Self::from_entries(
            TemplateOutput::Borrowed(output),
            load_entries_from_reader(Box::new(template))?,
        ))
    }

    /// Loads a path template and writes to an explicitly closeable stream.
    ///
    /// Keep a clone of `output` to observe Java-compatible close state after
    /// [`Self::finish`].
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn to_output_stream<W>(
        template: impl AsRef<Path>,
        output: ExcelOutputStream<W>,
    ) -> Result<Self>
    where
        W: Write + 'a,
    {
        Ok(Self::from_entries(
            TemplateOutput::Owned(Box::new(output)),
            load_entries(template.as_ref())?,
        ))
    }

    /// Loads a stream template and writes to an explicitly closeable stream.
    ///
    /// # Errors
    ///
    /// Returns an I/O or OOXML package error when the template cannot be read.
    pub fn from_reader_to_output_stream<R, W>(
        template: R,
        output: ExcelOutputStream<W>,
    ) -> Result<Self>
    where
        R: Read,
        W: Write + 'a,
    {
        Ok(Self::from_entries(
            TemplateOutput::Owned(Box::new(output)),
            load_entries_from_reader(Box::new(template))?,
        ))
    }

    /// Controls whether an owned output stream is closed by [`Self::finish`].
    ///
    /// The default is `true`, matching Java `EasyExcel`. Borrowed writers always
    /// remain caller-owned regardless of this setting.
    #[must_use]
    pub const fn auto_close_stream(mut self, enabled: bool) -> Self {
        self.auto_close_stream = enabled;
        self
    }

    /// Accumulates scalar `{key}` values for this workbook.
    ///
    /// Later fills replace earlier values for the same key, matching Java map
    /// filling before the workbook is finalized.
    ///
    /// # Errors
    ///
    /// Returns an error after the writer has finished.
    pub fn fill(&mut self, data: &TemplateData) -> Result<&mut Self> {
        self.fill_on_sheet(&TemplateSheet::first(), data)
    }

    /// Accumulates scalar `{key}` values for one selected worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error after the writer has finished.
    pub fn fill_on_sheet(
        &mut self,
        sheet: &TemplateSheet,
        data: &TemplateData,
    ) -> Result<&mut Self> {
        self.ensure_open()?;
        self.sheet_state_mut(sheet)
            .scalar
            .values
            .extend(data.values.clone());
        Ok(self)
    }

    /// Accumulates a named or unnamed collection fill.
    ///
    /// Repeated calls with the same prefix append rows. Java maintains one
    /// cursor per prefix; therefore changing the direction/configuration for an
    /// already-used prefix is rejected instead of silently restarting it.
    ///
    /// # Errors
    ///
    /// Returns an error after finish or when a prefix changes its fill config.
    pub fn fill_list(&mut self, data: &FillWrapper, config: FillConfig) -> Result<&mut Self> {
        self.fill_list_on_sheet(&TemplateSheet::first(), data, config)
    }

    /// Accumulates a collection fill for one selected worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error after finish or when a prefix changes its fill config.
    pub fn fill_list_on_sheet(
        &mut self,
        sheet: &TemplateSheet,
        data: &FillWrapper,
        config: FillConfig,
    ) -> Result<&mut Self> {
        self.ensure_open()?;
        if data.rows().is_empty() {
            return Ok(self);
        }
        let state = self.sheet_state_mut(sheet);
        if let Some(pending) = state
            .collections
            .iter_mut()
            .find(|pending| pending.wrapper.name == data.name)
        {
            if pending.config != config {
                return Err(ExcelError::Format(format!(
                    "collection fill prefix {:?} cannot change configuration between fills",
                    data.name()
                )));
            }
            pending.wrapper.rows.extend(data.rows.iter().cloned());
            return Ok(self);
        }
        state.collections.push(PendingCollectionFill {
            wrapper: data.clone(),
            config,
        });
        Ok(self)
    }

    /// Queues ordinary rows after the template fill cursor.
    ///
    /// This corresponds to Java's `excelWriter.write(rows, writeSheet)` after
    /// one or more `fill` calls. It is primarily intended for summary rows when
    /// the collection placeholder is the final template row.
    ///
    /// # Errors
    ///
    /// Returns an error after the writer has finished.
    pub fn write_rows(
        &mut self,
        rows: impl IntoIterator<Item = Vec<CellValue>>,
    ) -> Result<&mut Self> {
        self.write_rows_on_sheet(&TemplateSheet::first(), rows)
    }

    /// Queues ordinary rows after the fill cursor of one selected worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error after the writer has finished.
    pub fn write_rows_on_sheet(
        &mut self,
        sheet: &TemplateSheet,
        rows: impl IntoIterator<Item = Vec<CellValue>>,
    ) -> Result<&mut Self> {
        self.ensure_open()?;
        self.sheet_state_mut(sheet).appended_rows.extend(rows);
        Ok(self)
    }

    /// Writes the completed XLSX package. Repeated calls are no-ops.
    ///
    /// # Errors
    ///
    /// Returns an XML, ZIP, or output I/O error.
    pub fn finish(&mut self) -> Result<()> {
        if self.finished {
            return Ok(());
        }
        for sheet in self.resolved_sheet_fills()? {
            for pending in &sheet.collections {
                replace_collection_placeholders_in_sheet(
                    &mut self.entries,
                    &sheet.worksheet,
                    &pending.wrapper,
                    pending.config,
                );
            }
            replace_scalar_cells_in_sheet(&mut self.entries, &sheet.worksheet, &sheet.scalar)?;
            append_rows_to_sheet(&mut self.entries, &sheet.worksheet, &sheet.appended_rows)?;
        }
        self.finished = true;
        write_entries_to_output(&mut self.output, &self.entries, self.auto_close_stream)
    }

    /// Returns whether [`Self::finish`] has run.
    #[must_use]
    pub const fn is_finished(&self) -> bool {
        self.finished
    }

    fn ensure_open(&self) -> Result<()> {
        if self.finished {
            Err(ExcelError::Unsupported(
                "template writer already finished".to_owned(),
            ))
        } else {
            Ok(())
        }
    }

    fn sheet_state_mut(&mut self, sheet: &TemplateSheet) -> &mut PendingSheetFill {
        if let Some(index) = self
            .sheets
            .iter()
            .position(|pending| same_sheet(&pending.sheet, sheet))
        {
            return &mut self.sheets[index];
        }
        self.sheets.push(PendingSheetFill::new(sheet.clone()));
        self.sheets.last_mut().expect("sheet state was just pushed")
    }

    fn resolved_sheet_fills(&self) -> Result<Vec<ResolvedSheetFill>> {
        let mut resolved: Vec<ResolvedSheetFill> = Vec::new();
        for pending_sheet in &self.sheets {
            let worksheet = worksheet_path(&self.entries, &pending_sheet.sheet)?;
            if let Some(sheet) = resolved
                .iter_mut()
                .find(|sheet| sheet.worksheet.eq_ignore_ascii_case(&worksheet))
            {
                sheet
                    .scalar
                    .values
                    .extend(pending_sheet.scalar.values.clone());
                for pending_collection in &pending_sheet.collections {
                    if let Some(collection) = sheet.collections.iter_mut().find(|collection| {
                        collection.wrapper.name == pending_collection.wrapper.name
                    }) {
                        if collection.config != pending_collection.config {
                            return Err(ExcelError::Format(format!(
                                "collection fill prefix {:?} cannot change configuration between fills",
                                pending_collection.wrapper.name()
                            )));
                        }
                        collection
                            .wrapper
                            .rows
                            .extend(pending_collection.wrapper.rows.iter().cloned());
                    } else {
                        sheet.collections.push(pending_collection.clone());
                    }
                }
                sheet
                    .appended_rows
                    .extend(pending_sheet.appended_rows.iter().cloned());
            } else {
                resolved.push(ResolvedSheetFill {
                    worksheet,
                    scalar: pending_sheet.scalar.clone(),
                    collections: pending_sheet.collections.clone(),
                    appended_rows: pending_sheet.appended_rows.clone(),
                });
            }
        }
        Ok(resolved)
    }
}

fn same_sheet(left: &TemplateSheet, right: &TemplateSheet) -> bool {
    match (left, right) {
        (
            TemplateSheet::First | TemplateSheet::Index(0),
            TemplateSheet::First | TemplateSheet::Index(0),
        ) => true,
        (TemplateSheet::Index(left), TemplateSheet::Index(right)) => left == right,
        (TemplateSheet::Name(left), TemplateSheet::Name(right)) => left == right,
        _ => false,
    }
}

trait ReadSeek: Read + Seek {}

impl<T: Read + Seek> ReadSeek for T {}

trait WriteSeek: Write + Seek + Any {
    fn into_any(self: Box<Self>) -> Box<dyn Any>;
}

impl<T: Write + Seek + Any> WriteSeek for T {
    fn into_any(self: Box<Self>) -> Box<dyn Any> {
        self
    }
}

type ArchiveWriter = ZipWriter<Box<dyn WriteSeek>>;

/// Fills scalar `{key}` placeholders while preserving the XLSX package structure.
///
/// The template and output paths may be identical because the source archive is
/// fully loaded before the destination is opened.
///
/// # Errors
///
/// Returns an I/O or format error for invalid ZIP/OOXML input or output failures.
/// Legacy `.xls` templates are now supported via BIFF8 placeholder replacement.
pub fn fill_xlsx_template(template: &Path, output: &Path, data: &TemplateData) -> Result<()> {
    if template
        .extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls"))
    {
        return fill_xls_template_scalar(template, output, data);
    }
    let mut writer = ExcelTemplateWriter::new(template, output)?;
    writer.sheets[0].scalar.values.extend(data.values.clone());
    writer.finish()
}

/// Expands Java EasyExcel-style collection placeholders in an XLSX template.
///
/// Unnamed wrappers use `{.field}` while named wrappers use `{name.field}`.
///
/// # Errors
///
/// Returns an I/O or format error when the package cannot be read or written.
pub fn fill_xlsx_template_list(
    template: &Path,
    output: &Path,
    data: &FillWrapper,
    config: FillConfig,
) -> Result<()> {
    if template
        .extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls"))
    {
        return fill_xls_template_list(template, output, data, config);
    }
    let mut writer = ExcelTemplateWriter::new(template, output)?;
    if !data.rows().is_empty() {
        writer.sheets[0].collections.push(PendingCollectionFill {
            wrapper: data.clone(),
            config,
        });
    }
    writer.finish()
}

// ---------------------------------------------------------------------------
// BIFF8 (.xls) template fill — Phase 5
// ---------------------------------------------------------------------------

/// Replaces `{key}` placeholders in a legacy BIFF8 `.xls` template with
/// `TemplateData` scalar values. Mirrors Java's HSSFWorkbook-level fill
/// for XLS workbooks.
fn fill_xls_template_scalar(template: &Path, output: &Path, data: &TemplateData) -> Result<()> {
    let bytes = std::fs::read(template)?;
    let mut pkg = easyexcel_writer::biff8::Biff8TemplatePackage::from_bytes(&bytes)?;
    let placeholders = pkg.scan_placeholders();
    for (sheet_name, row, col, text) in &placeholders {
        let key = text.trim_start_matches('{').trim_end_matches('}').to_string();
        if let Some(value) = data.values.get(&key) {
            let replacement = value.as_text();
            pkg.replace_label(sheet_name, *row, *col, &replacement)?;
        }
    }
    pkg.save_to_path(output)
}

/// Replaces list placeholders in a BIFF8 `.xls` template.
fn fill_xls_template_list(
    template: &Path,
    output: &Path,
    data: &FillWrapper,
    _config: FillConfig,
) -> Result<()> {
    let bytes = std::fs::read(template)?;
    let mut pkg = easyexcel_writer::biff8::Biff8TemplatePackage::from_bytes(&bytes)?;
    let placeholders = pkg.scan_placeholders();
    let prefix = data.name().map(|n| format!("{n}.")).unwrap_or_default();
    let is_dot = prefix.is_empty();

    for (sheet_name, row, col, text) in &placeholders {
        let key = if is_dot && text.starts_with("{.") {
            text.trim_start_matches("{.").trim_end_matches('}').to_string()
        } else if !prefix.is_empty() && text.starts_with(&format!("{{{prefix}")) {
            text.trim_start_matches(&format!("{{{prefix}"))
                .trim_end_matches('}')
                .to_string()
        } else if text.starts_with('{') {
            text.trim_start_matches('{').trim_end_matches('}').to_string()
        } else {
            continue;
        };
        if key.is_empty() { continue; }
        for template_row in data.rows() {
            if let Some(value) = template_row.values.get(&key) {
                let replacement = value.as_text();
                pkg.replace_label(sheet_name, *row, *col, &replacement)?;
                break;
            }
        }
    }
    pkg.save_to_path(output)
}

#[cfg(test)]
fn replace_collection_placeholders(
    entries: &mut [TemplateEntry],
    wrapper: &FillWrapper,
    config: FillConfig,
) {
    replace_collection_placeholders_matching(entries, None, wrapper, config);
}

fn replace_collection_placeholders_in_sheet(
    entries: &mut [TemplateEntry],
    worksheet: &str,
    wrapper: &FillWrapper,
    config: FillConfig,
) {
    replace_collection_placeholders_matching(entries, Some(worksheet), wrapper, config);
}

fn replace_collection_placeholders_matching(
    entries: &mut [TemplateEntry],
    worksheet: Option<&str>,
    wrapper: &FillWrapper,
    config: FillConfig,
) {
    if wrapper.rows().is_empty() {
        return;
    }
    let shared_strings = entries
        .iter()
        .find(|entry| entry.name.eq_ignore_ascii_case("xl/sharedStrings.xml"))
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .map(shared_string_values)
        .unwrap_or_default();
    for entry in entries.iter_mut().filter(|entry| {
        worksheet.map_or_else(
            || entry.name.starts_with("xl/worksheets/"),
            |worksheet| entry.name.eq_ignore_ascii_case(worksheet),
        ) && Path::new(&entry.name)
            .extension()
            .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
    }) {
        let Ok(xml) = std::str::from_utf8(&entry.bytes) else {
            continue;
        };
        let expanded = match config.get_direction() {
            FillDirection::Vertical => expand_vertical_rows(xml, wrapper, config, &shared_strings),
            FillDirection::Horizontal => expand_horizontal_cells(xml, wrapper, &shared_strings),
        };
        if let Some(expanded) = expanded {
            entry.bytes = update_worksheet_dimension(&expanded).into_bytes();
            break;
        }
    }
}

fn shared_string_values(xml: &str) -> Vec<String> {
    let mut values = Vec::new();
    let mut remaining = xml;
    while let Some(start) = remaining.find("<si") {
        let item = &remaining[start..];
        let Some(open_end) = item.find('>') else {
            break;
        };
        let Some(close) = item.find("</si>") else {
            break;
        };
        values.push(text_node_values(&item[open_end + 1..close]));
        remaining = &item[close + 5..];
    }
    values
}

fn text_node_values(xml: &str) -> String {
    let mut value = String::new();
    let mut remaining = xml;
    while let Some(start) = remaining.find("<t") {
        let text = &remaining[start..];
        let Some(open_end) = text.find('>') else {
            break;
        };
        let Some(close) = text.find("</t>") else {
            break;
        };
        value.push_str(&text[open_end + 1..close]);
        remaining = &text[close + 4..];
    }
    value
}

fn expand_vertical_rows(
    xml: &str,
    wrapper: &FillWrapper,
    config: FillConfig,
    shared_strings: &[String],
) -> Option<String> {
    let (start, end, row, _, _, _) = find_collection_row(xml, wrapper, shared_strings)?;
    let first = fill_row_cells(
        row,
        wrapper.rows().first()?,
        wrapper.name(),
        shared_strings,
        config.get_auto_style(),
    );
    if config.get_force_new_row() {
        let template_row = row_index(row)?;
        let mut rows = first;
        for (offset, data) in wrapper.rows().iter().enumerate().skip(1) {
            rows.push_str(&collection_only_row(
                row,
                data,
                wrapper,
                shared_strings,
                config.get_auto_style(),
                offset,
            ));
        }
        let delta = wrapper.rows().len().saturating_sub(1);
        let suffix = shift_rows(&xml[end..], delta);
        let expanded = format!("{}{}{}", &xml[..start], rows, suffix);
        return Some(shift_worksheet_metadata(&expanded, template_row + 1, delta));
    }

    let template_row = row_index(row)?;
    let mut suffix = xml[end..].to_owned();
    for (offset, data) in wrapper.rows().iter().enumerate().skip(1) {
        let row = collection_only_row(
            row,
            data,
            wrapper,
            shared_strings,
            config.get_auto_style(),
            offset,
        );
        suffix = upsert_collection_row(&suffix, &row, template_row + offset);
    }
    Some(format!("{}{}{}", &xml[..start], first, suffix))
}

fn collection_only_row(
    template_row: &str,
    data: &TemplateData,
    wrapper: &FillWrapper,
    shared_strings: &[String],
    auto_style: bool,
    row_offset: usize,
) -> String {
    let Some(tag_end) = template_row.find('>') else {
        return template_row.to_owned();
    };
    let mut row = shift_row(&template_row[..=tag_end], row_offset, 0);
    for (_, _, cell) in collection_cells(template_row, wrapper, shared_strings) {
        let filled = fill_cell(cell, data, wrapper.name(), shared_strings, auto_style);
        row.push_str(&shift_row(&filled, row_offset, 0));
    }
    row.push_str("</row>");
    row
}

fn collection_cells<'a>(
    row: &'a str,
    wrapper: &FillWrapper,
    shared_strings: &[String],
) -> Vec<(usize, usize, &'a str)> {
    all_cells(row)
        .into_iter()
        .filter(|(_, _, cell)| {
            cell_value(cell, shared_strings)
                .is_some_and(|value| contains_collection_marker(&value, wrapper))
        })
        .collect()
}

fn upsert_collection_row(xml: &str, collection_row: &str, target_row: usize) -> String {
    let mut offset = 0;
    while let Some(relative_start) = xml[offset..].find("<row") {
        let start = offset + relative_start;
        let Some(relative_end) = xml[start..].find("</row>") else {
            break;
        };
        let end = start + relative_end + 6;
        let existing = &xml[start..end];
        match row_index(existing) {
            Some(row) if row == target_row => {
                let merged = merge_collection_cells(existing, collection_row);
                return format!("{}{}{}", &xml[..start], merged, &xml[end..]);
            }
            Some(row) if row > target_row => {
                return format!("{}{}{}", &xml[..start], collection_row, &xml[start..]);
            }
            _ => offset = end,
        }
    }
    if let Some(end) = xml.find("</sheetData>") {
        return format!("{}{}{}", &xml[..end], collection_row, &xml[end..]);
    }
    format!("{xml}{collection_row}")
}

fn merge_collection_cells(existing_row: &str, collection_row: &str) -> String {
    let mut merged = existing_row.to_owned();
    for (_, _, cell) in all_cells(collection_row) {
        let Some(reference) = attribute_value(cell, "r") else {
            continue;
        };
        if let Some((start, end, _)) = all_cells(&merged)
            .into_iter()
            .find(|(_, _, existing)| attribute_value(existing, "r") == Some(reference))
        {
            merged.replace_range(start..end, cell);
        } else if let Some(end) = merged.rfind("</row>") {
            merged.insert_str(end, cell);
        }
    }
    merged
}

fn all_cells(row: &str) -> Vec<(usize, usize, &str)> {
    let mut cells = Vec::new();
    let mut offset = 0;
    while let Some((start, end)) = find_next_cell(row, offset) {
        cells.push((start, end, &row[start..end]));
        offset = end;
    }
    cells
}

/// Finds the next OOXML cell element (`<c ...>` / `<c .../>`).
///
/// Must not match similarly-prefixed tags such as `<cols>` / `<col>` — those
/// false positives previously rewrote worksheet XML during scalar fill and
/// left `complex.xlsx` unreadable by quick_xml.
fn find_next_cell(xml: &str, from: usize) -> Option<(usize, usize)> {
    let bytes = xml.as_bytes();
    let mut search = from;
    while search < xml.len() {
        let relative = xml[search..].find("<c")?;
        let start = search + relative;
        let after_c = start + 2;
        let next = *bytes.get(after_c)?;
        // Local name must be exactly `c` (followed by space, `/`, or `>`).
        if !matches!(next, b' ' | b'\t' | b'\n' | b'\r' | b'/' | b'>') {
            search = after_c;
            continue;
        }
        let relative_gt = xml[after_c..].find('>')?;
        let gt = after_c + relative_gt;
        // Self-closing `<c .../>` — do not scan forward for `</c>`.
        if gt > start && bytes[gt - 1] == b'/' {
            return Some((start, gt + 1));
        }
        let relative_end = xml[gt..].find("</c>")?;
        let end = gt + relative_end + 4;
        return Some((start, end));
    }
    None
}

fn row_index(row: &str) -> Option<usize> {
    attribute_value(row, "r")?.parse().ok()
}

fn update_worksheet_dimension(xml: &str) -> String {
    let mut bounds: Option<(usize, usize, usize, usize)> = None;
    for (_, _, cell) in all_cells(xml) {
        let Some(reference) = attribute_value(cell, "r") else {
            continue;
        };
        let Some((column, row)) = parse_cell_reference(reference) else {
            continue;
        };
        bounds = Some(bounds.map_or((column, row, column, row), |current| {
            (
                current.0.min(column),
                current.1.min(row),
                current.2.max(column),
                current.3.max(row),
            )
        }));
    }
    let Some((first_column, first_row, last_column, last_row)) = bounds else {
        return xml.to_owned();
    };
    let reference = format!(
        "{}{}:{}{}",
        column_name(first_column),
        first_row,
        column_name(last_column),
        last_row
    );
    replace_tag_attribute(xml, "dimension", "ref", &reference)
}

fn shift_worksheet_metadata(xml: &str, threshold_row: usize, delta: usize) -> String {
    if delta == 0 {
        return xml.to_owned();
    }
    let mut shifted = xml.to_owned();
    for (tag, attribute) in [
        ("mergeCell", "ref"),
        ("hyperlink", "ref"),
        ("autoFilter", "ref"),
        ("dataValidation", "sqref"),
        ("conditionalFormatting", "sqref"),
    ] {
        shifted = shift_tag_references(&shifted, tag, attribute, threshold_row, delta);
    }
    shift_formula_elements(&shifted, threshold_row, delta)
}

fn shift_tag_references(
    xml: &str,
    tag: &str,
    attribute: &str,
    threshold_row: usize,
    delta: usize,
) -> String {
    let mut output = String::new();
    let mut offset = 0;
    let marker = format!("<{tag}");
    while let Some(relative_start) = xml[offset..].find(&marker) {
        let start = offset + relative_start;
        let Some(relative_end) = xml[start..].find('>') else {
            break;
        };
        let end = start + relative_end + 1;
        output.push_str(&xml[offset..start]);
        let element = &xml[start..end];
        let shifted = attribute_value(element, attribute).map_or_else(
            || element.to_owned(),
            |value| {
                replace_attribute(
                    element,
                    attribute,
                    &shift_reference_list(value, threshold_row, delta),
                )
            },
        );
        output.push_str(&shifted);
        offset = end;
    }
    output.push_str(&xml[offset..]);
    output
}

fn shift_reference_list(value: &str, threshold_row: usize, delta: usize) -> String {
    value
        .split_whitespace()
        .map(|range| {
            range
                .split(':')
                .map(|reference| shift_a1_reference(reference, threshold_row, delta))
                .collect::<Vec<_>>()
                .join(":")
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn shift_formula_elements(xml: &str, threshold_row: usize, delta: usize) -> String {
    let mut output = String::new();
    let mut offset = 0;
    while let Some(relative_start) = xml[offset..].find("<f") {
        let start = offset + relative_start;
        let Some(open_end) = xml[start..].find('>') else {
            break;
        };
        let content_start = start + open_end + 1;
        let Some(relative_end) = xml[content_start..].find("</f>") else {
            break;
        };
        let content_end = content_start + relative_end;
        output.push_str(&xml[offset..content_start]);
        output.push_str(&shift_formula_references(
            &xml[content_start..content_end],
            threshold_row,
            delta,
        ));
        offset = content_end;
    }
    output.push_str(&xml[offset..]);
    output
}

fn shift_formula_references(formula: &str, threshold_row: usize, delta: usize) -> String {
    let bytes = formula.as_bytes();
    let mut output = String::new();
    let mut offset = 0;
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'$' && !bytes[index].is_ascii_alphabetic() {
            index += 1;
            continue;
        }
        let start = index;
        if bytes[index] == b'$' {
            index += 1;
        }
        let column_start = index;
        while index < bytes.len() && bytes[index].is_ascii_alphabetic() {
            index += 1;
        }
        let column_end = index;
        if index < bytes.len() && bytes[index] == b'$' {
            index += 1;
        }
        let row_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        let valid = column_end > column_start
            && row_start < index
            && column_end - column_start <= 3
            && (start == 0 || !is_formula_identifier(bytes[start - 1]))
            && (index == bytes.len()
                || (!is_formula_identifier(bytes[index])
                    && bytes[index] != b'!'
                    && bytes[index] != b'('));
        if valid {
            output.push_str(&formula[offset..start]);
            output.push_str(&shift_a1_reference(
                &formula[start..index],
                threshold_row,
                delta,
            ));
            offset = index;
        }
    }
    output.push_str(&formula[offset..]);
    output
}

const fn is_formula_identifier(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_'
}

fn shift_a1_reference(reference: &str, threshold_row: usize, delta: usize) -> String {
    let Some((column, row)) = parse_cell_reference(reference) else {
        return reference.to_owned();
    };
    if row < threshold_row {
        return reference.to_owned();
    }
    let absolute_column = reference.starts_with('$');
    let row_marker = reference
        .bytes()
        .position(|byte| byte.is_ascii_digit())
        .is_some_and(|index| index > 0 && reference.as_bytes()[index - 1] == b'$');
    format!(
        "{}{}{}{}",
        if absolute_column { "$" } else { "" },
        column_name(column),
        if row_marker { "$" } else { "" },
        row + delta
    )
}

fn parse_cell_reference(reference: &str) -> Option<(usize, usize)> {
    let normalized = reference.replace('$', "");
    let split = normalized.bytes().position(|byte| byte.is_ascii_digit())?;
    if split == 0
        || !normalized[..split]
            .bytes()
            .all(|byte| byte.is_ascii_alphabetic())
    {
        return None;
    }
    let column = normalized[..split]
        .bytes()
        .try_fold(0_usize, |value, byte| {
            value
                .checked_mul(26)?
                .checked_add(usize::from(byte.to_ascii_uppercase() - b'A' + 1))
        })?;
    if column > 16_384 {
        return None;
    }
    let row = normalized[split..].parse::<usize>().ok()?;
    (row > 0).then_some((column, row))
}

fn replace_tag_attribute(xml: &str, tag: &str, attribute: &str, value: &str) -> String {
    let marker = format!("<{tag}");
    let Some(start) = xml.find(&marker) else {
        return xml.to_owned();
    };
    let Some(relative_end) = xml[start..].find('>') else {
        return xml.to_owned();
    };
    let end = start + relative_end + 1;
    let replaced = replace_attribute(&xml[start..end], attribute, value);
    format!("{}{}{}", &xml[..start], replaced, &xml[end..])
}

fn expand_horizontal_cells(
    xml: &str,
    wrapper: &FillWrapper,
    shared_strings: &[String],
) -> Option<String> {
    let mut output = String::with_capacity(xml.len());
    let mut offset = 0;
    let mut changed = false;
    while let Some(relative_start) = xml[offset..].find("<row") {
        let start = offset + relative_start;
        let Some(relative_end) = xml[start..].find("</row>") else {
            break;
        };
        let end = start + relative_end + 6;
        output.push_str(&xml[offset..start]);
        let row = &xml[start..end];
        let cells = collection_cells(row, wrapper, shared_strings);
        if cells.is_empty() {
            output.push_str(row);
        } else {
            changed = true;
            let mut cell_offset = 0;
            for (cell_start, cell_end, cell) in cells {
                output.push_str(&row[cell_offset..cell_start]);
                for (column_offset, data) in wrapper.rows().iter().enumerate() {
                    let filled = fill_cell(cell, data, wrapper.name(), shared_strings, true);
                    output.push_str(&shift_row(&filled, 0, column_offset));
                }
                cell_offset = cell_end;
            }
            output.push_str(&row[cell_offset..]);
        }
        offset = end;
    }
    output.push_str(&xml[offset..]);
    changed.then_some(output)
}

fn find_collection_row<'a>(
    xml: &'a str,
    wrapper: &FillWrapper,
    shared_strings: &[String],
) -> Option<(usize, usize, &'a str, usize, usize, &'a str)> {
    let mut offset = 0;
    while let Some(relative_start) = xml[offset..].find("<row") {
        let start = offset + relative_start;
        let end = start + xml[start..].find("</row>")? + 6;
        let row = &xml[start..end];
        if let Some((cell_start, cell_end, cell)) =
            find_collection_cell(row, wrapper, shared_strings)
        {
            return Some((start, end, row, cell_start, cell_end, cell));
        }
        offset = end;
    }
    None
}

fn find_collection_cell<'a>(
    row: &'a str,
    wrapper: &FillWrapper,
    shared_strings: &[String],
) -> Option<(usize, usize, &'a str)> {
    let mut offset = 0;
    while let Some((start, end)) = find_next_cell(row, offset) {
        let cell = &row[start..end];
        if cell_value(cell, shared_strings)
            .is_some_and(|value| contains_collection_marker(&value, wrapper))
        {
            return Some((start, end, cell));
        }
        offset = end;
    }
    None
}

fn fill_row_cells(
    row: &str,
    data: &TemplateData,
    prefix: Option<&str>,
    shared_strings: &[String],
    auto_style: bool,
) -> String {
    let mut output = String::new();
    let mut offset = 0;
    while let Some((start, end)) = find_next_cell(row, offset) {
        output.push_str(&row[offset..start]);
        output.push_str(&fill_cell(
            &row[start..end],
            data,
            prefix,
            shared_strings,
            auto_style,
        ));
        offset = end;
    }
    output.push_str(&row[offset..]);
    output
}

fn fill_cell(
    cell: &str,
    data: &TemplateData,
    prefix: Option<&str>,
    shared_strings: &[String],
    auto_style: bool,
) -> String {
    let Some(tag_end) = cell.find('>') else {
        return cell.to_owned();
    };
    let Some(value) = cell_value(cell, shared_strings) else {
        return cell.to_owned();
    };
    if let Some(typed_value) = exact_collection_value(&value, data, prefix) {
        return render_typed_cell(cell, typed_value, auto_style);
    }
    let filled = replace_collection_values(&value, data, prefix);
    if filled == value {
        return cell.to_owned();
    }
    let mut start = cell[..=tag_end].replace(" t=\"s\"", "");
    if !auto_style {
        start = remove_attribute(&start, "s");
    }
    if start.contains(" t=\"") {
        start = replace_attribute(&start, "t", "inlineStr");
    } else {
        start.insert_str(start.len() - 1, " t=\"inlineStr\"");
    }
    format!("{start}<is><t>{}</t></is></c>", escape_xml(&filled))
}

fn exact_collection_value<'a>(
    placeholder: &str,
    data: &'a TemplateData,
    prefix: Option<&str>,
) -> Option<&'a CellValue> {
    let variable = placeholder.strip_prefix('{')?.strip_suffix('}')?;
    let key = match prefix {
        Some(prefix) => variable.strip_prefix(prefix)?.strip_prefix('.')?,
        None => variable.strip_prefix('.')?,
    };
    data.values().get(key)
}

fn exact_scalar_value<'a>(placeholder: &str, data: &'a TemplateData) -> Option<&'a CellValue> {
    let key = placeholder.strip_prefix('{')?.strip_suffix('}')?;
    (!key.starts_with('.') && !key.ends_with('.'))
        .then(|| data.values().get(key))
        .flatten()
}

fn render_typed_cell(cell: &str, value: &CellValue, auto_style: bool) -> String {
    let Some(tag_end) = cell.find('>') else {
        return cell.to_owned();
    };
    let mut start = cell[..=tag_end].to_owned();
    if !auto_style {
        start = remove_attribute(&start, "s");
    }
    start = remove_attribute(&start, "t");
    match value {
        CellValue::Empty | CellValue::Image(_) => format!("{start}</c>"),
        CellValue::String(value) | CellValue::Hyperlink { text: value, .. } => {
            insert_cell_type(&mut start, "inlineStr");
            format!("{start}<is><t>{}</t></is></c>", escape_xml(value))
        }
        CellValue::Bool(value) => {
            insert_cell_type(&mut start, "b");
            format!("{start}<v>{}</v></c>", u8::from(*value))
        }
        CellValue::Int(value) => format!("{start}<v>{value}</v></c>"),
        CellValue::Float(value) => format!("{start}<v>{value}</v></c>"),
        CellValue::Decimal(value) => format!("{start}<v>{value}</v></c>"),
        CellValue::Date(value) => {
            insert_cell_type(&mut start, "d");
            format!("{start}<v>{}</v></c>", value.format("%Y-%m-%d"))
        }
        CellValue::DateTime(value) => {
            insert_cell_type(&mut start, "d");
            format!("{start}<v>{}</v></c>", value.format("%Y-%m-%dT%H:%M:%S"))
        }
        CellValue::Error(value) => {
            insert_cell_type(&mut start, "e");
            format!("{start}<v>{}</v></c>", escape_xml(value))
        }
        CellValue::Formula(value) => {
            format!("{start}<f>{}</f><v></v></c>", escape_xml(value))
        }
        CellValue::RichText(value) => {
            insert_cell_type(&mut start, "inlineStr");
            format!(
                "{start}<is><t>{}</t></is></c>",
                escape_xml(value.text_string())
            )
        }
        CellValue::Comment { value, .. } | CellValue::Images { value, .. } => {
            render_typed_cell(cell, value, auto_style)
        }
    }
}

fn insert_cell_type(start: &mut String, cell_type: &str) {
    start.insert_str(start.len() - 1, &format!(" t=\"{cell_type}\""));
}

fn cell_value(cell: &str, shared_strings: &[String]) -> Option<String> {
    if attribute_value(cell, "t") == Some("s") {
        let index = element_value(cell, "v")?.parse::<usize>().ok()?;
        return shared_strings.get(index).cloned();
    }
    let value = text_node_values(cell);
    (!value.is_empty()).then_some(value)
}

fn contains_collection_marker(value: &str, wrapper: &FillWrapper) -> bool {
    let prefix = wrapper
        .name()
        .map_or(".".to_owned(), |name| format!("{name}."));
    contains_unescaped(value, &format!("{{{prefix}"))
}

fn replace_collection_values(value: &str, data: &TemplateData, prefix: Option<&str>) -> String {
    replace_template_values(value, data.values(), prefix, false, false)
}

#[cfg(test)]
fn replace_scalar_cells(entries: &mut [TemplateEntry], data: &TemplateData) -> Result<()> {
    replace_scalar_cells_matching(entries, None, data)
}

fn replace_scalar_cells_in_sheet(
    entries: &mut [TemplateEntry],
    worksheet: &str,
    data: &TemplateData,
) -> Result<()> {
    replace_scalar_cells_matching(entries, Some(worksheet), data)
}

fn replace_scalar_cells_matching(
    entries: &mut [TemplateEntry],
    worksheet: Option<&str>,
    data: &TemplateData,
) -> Result<()> {
    let shared_strings = entries
        .iter()
        .find(|entry| entry.name == "xl/sharedStrings.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .map_or_else(Vec::new, shared_string_values);
    for entry in entries.iter_mut().filter(|entry| {
        !entry.is_dir
            && worksheet.map_or_else(
                || entry.name.starts_with("xl/worksheets/"),
                |worksheet| entry.name.eq_ignore_ascii_case(worksheet),
            )
            && Path::new(&entry.name)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
    }) {
        let xml = String::from_utf8(std::mem::take(&mut entry.bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        entry.bytes = replace_scalar_cells_in_xml(&xml, data, &shared_strings).into_bytes();
    }
    Ok(())
}

fn replace_scalar_cells_in_xml(
    xml: &str,
    data: &TemplateData,
    shared_strings: &[String],
) -> String {
    let mut output = String::with_capacity(xml.len());
    let mut offset = 0;
    while let Some((start, end)) = find_next_cell(xml, offset) {
        let cell = &xml[start..end];
        output.push_str(&xml[offset..start]);
        let replacement = cell_value(cell, shared_strings).map_or_else(
            || cell.to_owned(),
            |placeholder| {
                if let Some(value) = exact_scalar_value(&placeholder, data) {
                    return render_typed_cell(cell, value, true);
                }
                let filled =
                    replace_template_values(&placeholder, data.values(), None, true, false);
                if filled == placeholder {
                    cell.to_owned()
                } else {
                    render_typed_cell(cell, &CellValue::String(filled), true)
                }
            },
        );
        output.push_str(&replacement);
        offset = end;
    }
    output.push_str(&xml[offset..]);
    output
}

#[cfg(test)]
fn append_rows_to_first_sheet(
    entries: &mut [TemplateEntry],
    rows: &[Vec<CellValue>],
) -> Result<()> {
    append_rows_to_sheet(entries, "xl/worksheets/sheet1.xml", rows)
}

fn append_rows_to_sheet(
    entries: &mut [TemplateEntry],
    worksheet: &str,
    rows: &[Vec<CellValue>],
) -> Result<()> {
    if rows.is_empty() {
        return Ok(());
    }
    let Some(entry) = entries
        .iter_mut()
        .find(|entry| entry.name.eq_ignore_ascii_case(worksheet))
    else {
        return Err(ExcelError::Format(format!(
            "template does not contain {worksheet}"
        )));
    };
    let xml = String::from_utf8(std::mem::take(&mut entry.bytes))
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    entry.bytes = append_rows_to_xml(&xml, rows)?.into_bytes();
    Ok(())
}

fn append_rows_to_xml(xml: &str, rows: &[Vec<CellValue>]) -> Result<String> {
    let Some(sheet_data_end) = xml.find("</sheetData>") else {
        return Err(ExcelError::Format(
            "worksheet does not contain sheetData".to_owned(),
        ));
    };
    let next_row = worksheet_max_row(&xml[..sheet_data_end]).saturating_add(1);
    let mut appended = String::new();
    for (row_offset, values) in rows.iter().enumerate() {
        let row_index = next_row + row_offset;
        write!(appended, "<row r=\"{row_index}\">").expect("writing to String cannot fail");
        for (column, value) in values.iter().enumerate() {
            let reference = format!("{}{row_index}", column_name(column + 1));
            appended.push_str(&render_typed_cell(
                &format!("<c r=\"{reference}\"></c>"),
                value,
                true,
            ));
        }
        appended.push_str("</row>");
    }
    let expanded = format!(
        "{}{}{}",
        &xml[..sheet_data_end],
        appended,
        &xml[sheet_data_end..]
    );
    Ok(update_worksheet_dimension(&expanded))
}

fn worksheet_max_row(xml: &str) -> usize {
    let mut maximum = 0;
    let mut offset = 0;
    while let Some(relative_start) = xml[offset..].find("<row") {
        let start = offset + relative_start;
        let Some(relative_end) = xml[start..].find('>') else {
            break;
        };
        let end = start + relative_end + 1;
        if let Some(row) = row_index(&xml[start..end]) {
            maximum = maximum.max(row);
        }
        offset = end;
    }
    maximum
}

fn element_value<'a>(xml: &'a str, element: &str) -> Option<&'a str> {
    let start_marker = format!("<{element}>");
    let end_marker = format!("</{element}>");
    let start = xml.find(&start_marker)? + start_marker.len();
    let end = start + xml[start..].find(&end_marker)?;
    Some(&xml[start..end])
}

fn attribute_value<'a>(xml: &'a str, attribute: &str) -> Option<&'a str> {
    let marker = format!(" {attribute}=\"");
    let start = xml.find(&marker)? + marker.len();
    let end = start + xml[start..].find('"')?;
    Some(&xml[start..end])
}

fn replace_attribute(xml: &str, attribute: &str, value: &str) -> String {
    let Some(current) = attribute_value(xml, attribute) else {
        return xml.to_owned();
    };
    xml.replacen(
        &format!(" {attribute}=\"{current}\""),
        &format!(" {attribute}=\"{value}\""),
        1,
    )
}

fn remove_attribute(xml: &str, attribute: &str) -> String {
    let Some(current) = attribute_value(xml, attribute) else {
        return xml.to_owned();
    };
    xml.replacen(&format!(" {attribute}=\"{current}\""), "", 1)
}

fn shift_rows(xml: &str, delta: usize) -> String {
    if delta == 0 {
        return xml.to_owned();
    }
    let mut output = String::new();
    let mut offset = 0;
    while let Some(relative_start) = xml[offset..].find("<row") {
        let start = offset + relative_start;
        let Some(relative_end) = xml[start..].find("</row>") else {
            break;
        };
        let end = start + relative_end + 6;
        output.push_str(&xml[offset..start]);
        output.push_str(&shift_row(&xml[start..end], delta, 0));
        offset = end;
    }
    output.push_str(&xml[offset..]);
    output
}

fn shift_row(xml: &str, row_delta: usize, column_delta: usize) -> String {
    let mut shifted = xml.to_owned();
    let references = cell_references(xml);
    for reference in references.into_iter().rev() {
        let replacement = shift_cell_reference(reference.2, row_delta, column_delta);
        shifted.replace_range(reference.0..reference.1, &replacement);
    }
    if xml.starts_with("<row")
        && let Some(row) = attribute_value(xml, "r").and_then(|value| value.parse::<usize>().ok())
    {
        shifted = replace_attribute(&shifted, "r", &(row + row_delta).to_string());
    }
    shifted
}

fn cell_references(xml: &str) -> Vec<(usize, usize, &str)> {
    let mut references = Vec::new();
    let mut offset = 0;
    while let Some(relative) = xml[offset..].find(" r=\"") {
        let start = offset + relative + 4;
        let Some(length) = xml[start..].find('"') else {
            break;
        };
        let end = start + length;
        let value = &xml[start..end];
        if value.bytes().any(|byte| byte.is_ascii_alphabetic())
            && value.bytes().any(|byte| byte.is_ascii_digit())
        {
            references.push((start, end, value));
        }
        offset = end + 1;
    }
    references
}

fn shift_cell_reference(reference: &str, row_delta: usize, column_delta: usize) -> String {
    let split = reference
        .bytes()
        .position(|byte| byte.is_ascii_digit())
        .unwrap_or(reference.len());
    if split == 0
        || split == reference.len()
        || !reference[..split]
            .bytes()
            .all(|byte| byte.is_ascii_alphabetic())
    {
        return reference.to_owned();
    }
    let column = reference[..split].bytes().fold(0_usize, |value, byte| {
        value * 26 + usize::from(byte.to_ascii_uppercase() - b'A' + 1)
    });
    let Ok(row) = reference[split..].parse::<usize>() else {
        return reference.to_owned();
    };
    let row = row + row_delta;
    format!("{}{}", column_name(column + column_delta), row)
}

fn column_name(mut column: usize) -> String {
    let mut name = String::new();
    while column > 0 {
        column -= 1;
        name.insert(0, char::from(b'A' + u8::try_from(column % 26).unwrap_or(0)));
        column /= 26;
    }
    name
}

fn load_entries(path: &Path) -> Result<Vec<TemplateEntry>> {
    // Scalar `.xls` fill is handled by [`fill_xlsx_template`] before ZIP load.
    // Stateful ExcelTemplateWriter / collection fill stay OOXML-only.
    if path
        .extension()
        .and_then(|ext| ext.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls"))
    {
        return Err(ExcelError::Unsupported(
            // Java: ExcelWriter.fill on HSSFWorkbook. Rust fill is OOXML-only;
            // use with_template + doWrite (Biff8TemplatePackage) for .xls cells.
            "legacy XLS template fill is not supported".to_owned(),
        ));
    }
    load_entries_from(Box::new(File::open(path)?))
}

fn load_entries_from_reader(mut reader: Box<dyn Read + '_>) -> Result<Vec<TemplateEntry>> {
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    load_entries_from(Box::new(Cursor::new(bytes)))
}

fn load_entries_from(reader: Box<dyn ReadSeek>) -> Result<Vec<TemplateEntry>> {
    let mut archive = ZipArchive::new(reader).map_err(format_error)?;
    let mut entries = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).map_err(format_error)?;
        let mut bytes = Vec::new();
        if !entry.is_dir() {
            entry.read_to_end(&mut bytes)?;
        }
        entries.push(TemplateEntry {
            name: entry.name().to_owned(),
            is_dir: entry.is_dir(),
            compression: entry.compression(),
            unix_mode: entry.unix_mode(),
            bytes,
        });
    }
    Ok(entries)
}

#[cfg(test)]
fn replace_placeholders(xml: &str, values: &BTreeMap<String, CellValue>) -> String {
    replace_template_values(xml, values, None, true, true)
}

fn replace_template_values(
    input: &str,
    values: &BTreeMap<String, CellValue>,
    collection_prefix: Option<&str>,
    scalar_values: bool,
    escape_values: bool,
) -> String {
    let bytes = input.as_bytes();
    let mut output = String::with_capacity(input.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'\\'
            && bytes
                .get(index + 1)
                .is_some_and(|next| matches!(next, b'{' | b'}'))
        {
            output.push(char::from(bytes[index + 1]));
            index += 2;
            continue;
        }
        if bytes[index] == b'{'
            && let Some(relative_end) = input[index + 1..].find('}')
        {
            let end = index + relative_end + 1;
            let placeholder = &input[index + 1..end];
            let key = if scalar_values {
                Some(placeholder)
            } else {
                match collection_prefix {
                    Some(prefix) => placeholder
                        .strip_prefix(prefix)
                        .and_then(|value| value.strip_prefix('.')),
                    None => placeholder.strip_prefix('.'),
                }
            };
            if let Some(value) = key.and_then(|key| values.get(key)) {
                let value = value.as_text();
                if escape_values {
                    output.push_str(&escape_xml(&value));
                } else {
                    output.push_str(&value);
                }
                index = end + 1;
                continue;
            }
        }
        let character = input[index..]
            .chars()
            .next()
            .expect("index always points to a character boundary");
        output.push(character);
        index += character.len_utf8();
    }
    output
}

fn contains_unescaped(value: &str, marker: &str) -> bool {
    let mut offset = 0;
    while let Some(relative) = value[offset..].find(marker) {
        let index = offset + relative;
        let backslashes = value[..index]
            .bytes()
            .rev()
            .take_while(|byte| *byte == b'\\')
            .count();
        if backslashes % 2 == 0 {
            return true;
        }
        offset = index + marker.len();
    }
    false
}

fn escape_xml(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            other => escaped.push(other),
        }
    }
    escaped
}

fn worksheet_path(entries: &[TemplateEntry], sheet: &TemplateSheet) -> Result<String> {
    let workbook = entries
        .iter()
        .find(|entry| entry.name.eq_ignore_ascii_case("xl/workbook.xml"));
    let relationships = entries.iter().find(|entry| {
        entry
            .name
            .eq_ignore_ascii_case("xl/_rels/workbook.xml.rels")
    });
    if let (Some(workbook), Some(relationships)) = (workbook, relationships) {
        let workbook = std::str::from_utf8(&workbook.bytes)
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let relationships = std::str::from_utf8(&relationships.bytes)
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let sheets = workbook_sheets(workbook);
        let selected = match sheet {
            TemplateSheet::First => sheets.first(),
            TemplateSheet::Index(index) => sheets.get(*index),
            TemplateSheet::Name(name) => sheets.iter().find(|(sheet_name, _)| sheet_name == name),
        }
        .ok_or_else(|| ExcelError::SheetNotFound(template_sheet_label(sheet)))?;
        let target = workbook_relationship_target(relationships, &selected.1).ok_or_else(|| {
            ExcelError::Format(format!(
                "workbook relationship {} for sheet {} is missing",
                selected.1, selected.0
            ))
        })?;
        let normalized = normalize_workbook_target(target)?;
        return entries
            .iter()
            .find(|entry| entry.name.eq_ignore_ascii_case(&normalized))
            .map(|entry| entry.name.clone())
            .ok_or_else(|| {
                ExcelError::Format(format!(
                    "worksheet part {normalized} for sheet {} is missing",
                    selected.0
                ))
            });
    }

    let worksheets = entries
        .iter()
        .filter(|entry| {
            entry.name.starts_with("xl/worksheets/")
                && Path::new(&entry.name)
                    .extension()
                    .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
        })
        .collect::<Vec<_>>();
    let index = match sheet {
        TemplateSheet::First => 0,
        TemplateSheet::Index(index) => *index,
        TemplateSheet::Name(name) => {
            return Err(ExcelError::SheetNotFound(name.clone()));
        }
    };
    worksheets
        .get(index)
        .map(|entry| entry.name.clone())
        .ok_or_else(|| ExcelError::SheetNotFound(template_sheet_label(sheet)))
}

fn workbook_sheets(xml: &str) -> Vec<(String, String)> {
    xml_elements(xml, "sheet")
        .filter_map(|element| {
            Some((
                attribute_value(element, "name")?.to_owned(),
                attribute_value(element, "r:id")?.to_owned(),
            ))
        })
        .collect()
}

fn workbook_relationship_target<'a>(xml: &'a str, relationship_id: &str) -> Option<&'a str> {
    xml_elements(xml, "Relationship")
        .find(|element| attribute_value(element, "Id") == Some(relationship_id))
        .and_then(|element| attribute_value(element, "Target"))
}

fn xml_elements<'a>(xml: &'a str, name: &'a str) -> impl Iterator<Item = &'a str> {
    let marker = format!("<{name}");
    let mut offset = 0;
    std::iter::from_fn(move || {
        while let Some(relative_start) = xml[offset..].find(&marker) {
            let start = offset + relative_start;
            let after_name = start + marker.len();
            if xml
                .as_bytes()
                .get(after_name)
                .is_some_and(u8::is_ascii_alphanumeric)
            {
                offset = after_name;
                continue;
            }
            let end = start + xml[start..].find('>')? + 1;
            offset = end;
            return Some(&xml[start..end]);
        }
        None
    })
}

fn normalize_workbook_target(target: &str) -> Result<String> {
    let candidate = target
        .strip_prefix('/')
        .map_or_else(|| format!("xl/{target}"), str::to_owned);
    let mut components = Vec::new();
    for component in candidate.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                if components.pop().is_none() {
                    return Err(ExcelError::Format(format!(
                        "worksheet target escapes package root: {target}"
                    )));
                }
            }
            component => components.push(component),
        }
    }
    if components.is_empty() {
        return Err(ExcelError::Format(format!(
            "worksheet target is empty: {target}"
        )));
    }
    Ok(components.join("/"))
}

fn template_sheet_label(sheet: &TemplateSheet) -> String {
    match sheet {
        TemplateSheet::First => "0".to_owned(),
        TemplateSheet::Index(index) => index.to_string(),
        TemplateSheet::Name(name) => name.clone(),
    }
}

fn write_entries(path: &Path, entries: &[TemplateEntry]) -> Result<()> {
    match File::create(path) {
        Ok(writer) => write_file_entries(writer, entries),
        Err(error) => Err(error.into()),
    }
}

fn write_entries_to_output(
    output: &mut TemplateOutput<'_>,
    entries: &[TemplateEntry],
    auto_close_stream: bool,
) -> Result<()> {
    match output {
        TemplateOutput::Path(path) => write_entries(path, entries),
        TemplateOutput::Borrowed(writer) => {
            let bytes = encode_entries(entries)?;
            writer.write_all(&bytes)?;
            writer.flush()?;
            Ok(())
        }
        TemplateOutput::Owned(writer) => {
            let write_result = encode_entries(entries).and_then(|bytes| {
                writer
                    .write_all(&bytes)
                    .and_then(|()| writer.flush())
                    .map_err(ExcelError::from)
            });
            let close_result = if auto_close_stream {
                writer.close()
            } else {
                Ok(())
            };
            close_result.map_err(ExcelError::from)?;
            write_result
        }
    }
}

fn encode_entries(entries: &[TemplateEntry]) -> Result<Vec<u8>> {
    let writer = write_entries_to(Box::new(Cursor::new(Vec::new())), entries)?;
    archive_output_bytes(writer)
}

fn archive_output_bytes(writer: Box<dyn WriteSeek>) -> Result<Vec<u8>> {
    writer
        .into_any()
        .downcast::<Cursor<Vec<u8>>>()
        .map(|cursor| cursor.into_inner())
        .map_err(|_| ExcelError::Format("ZIP output buffer type changed".to_owned()))
}

fn write_file_entries(writer: File, entries: &[TemplateEntry]) -> Result<()> {
    let _ = write_entries_to(Box::new(writer), entries)?;
    Ok(())
}

fn write_entries_to(
    writer: Box<dyn WriteSeek>,
    entries: &[TemplateEntry],
) -> Result<Box<dyn WriteSeek>> {
    let mut writer = Some(ZipWriter::new(writer));
    for entry in entries {
        let mut options = SimpleFileOptions::default().compression_method(entry.compression);
        if let Some(mode) = entry.unix_mode {
            options = options.unix_permissions(mode);
        }
        if entry.is_dir {
            let mut operation = |writer: &mut ArchiveWriter| {
                writer
                    .add_directory(&entry.name, options)
                    .map_err(format_error)
            };
            zip_writer_operation(&mut writer, &mut operation)?;
        } else {
            let mut start = |writer: &mut ArchiveWriter| {
                writer
                    .start_file(&entry.name, options)
                    .map_err(format_error)
            };
            zip_writer_operation(&mut writer, &mut start)?;
            let mut write = |writer: &mut ArchiveWriter| {
                writer.write_all(&entry.bytes).map_err(ExcelError::from)
            };
            zip_writer_operation(&mut writer, &mut write)?;
        }
    }
    finish_zip_writer(&mut writer)
}

fn finish_zip_writer(writer: &mut Option<ArchiveWriter>) -> Result<Box<dyn WriteSeek>> {
    let Some(writer) = writer.take() else {
        return Err(ExcelError::Format("ZIP writer is unavailable".to_owned()));
    };
    match catch_unwind(AssertUnwindSafe(|| writer.finish())) {
        Ok(Ok(output)) => Ok(output),
        Ok(Err(error)) => Err(format_error(error)),
        Err(_) => Err(ExcelError::Format(
            "ZIP writer panicked while finalizing output".to_owned(),
        )),
    }
}

fn zip_writer_operation(
    writer: &mut Option<ArchiveWriter>,
    operation: &mut dyn FnMut(&mut ArchiveWriter) -> Result<()>,
) -> Result<()> {
    let Some(active) = writer.as_mut() else {
        return Err(ExcelError::Format("ZIP writer is unavailable".to_owned()));
    };
    match catch_unwind(AssertUnwindSafe(|| operation(active))) {
        Ok(Ok(())) => Ok(()),
        Ok(Err(error)) => {
            let damaged = writer.take().expect("active writer exists");
            std::mem::forget(damaged);
            Err(error)
        }
        Err(_) => {
            let damaged = writer.take().expect("active writer exists");
            std::mem::forget(damaged);
            Err(ExcelError::Format(
                "ZIP writer panicked while processing output".to_owned(),
            ))
        }
    }
}

fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[cfg(test)]
mod tests;
