//! XLSX, XLS, and CSV readers backed by Calamine and the Rust CSV engine.

use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{BufRead, BufReader, Cursor, Read, Seek, SeekFrom};
use std::path::Path;
use std::sync::Arc;

use calamine::{
    Data, DataRef, ExcelDateTime, ExcelDateTimeType, Range, Reader, Xls, Xlsx, open_workbook,
};
use easyexcel_core::{
    AnalysisContext, CellExtra, CellExtraType, CellValue, ConverterRegistry, CsvCharset,
    CustomReadObject, ErrorAction, ExcelError, ExcelRow, FormulaData, ReadDefaultReturn,
    ReadListener, Result, RowData,
};
use encoding_rs::Encoding;
use encoding_rs_io::DecodeReaderBytesBuilder;

mod locale;
mod locale_generated;
mod xlsx_rows;

pub use locale::ExcelLocale;

use xlsx_rows::{XlsxDisplayCellReader, XlsxRowMetadata};

/// Selects a worksheet by index, name, or all sheets.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum SheetSelector {
    /// The first worksheet.
    #[default]
    First,
    /// A zero-based worksheet index.
    Index(usize),
    /// A worksheet name.
    Name(String),
    /// Every worksheet in workbook order.
    All,
}

/// Controls how General-format extreme numbers are displayed while reading.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ScientificFormatMode {
    /// Match Java `EasyExcel`'s default and avoid scientific notation.
    #[default]
    Plain,
    /// Use Java `EasyExcel`'s `0.#####E0` scientific representation.
    Scientific,
}

impl ScientificFormatMode {
    const fn is_enabled(self) -> bool {
        matches!(self, Self::Scientific)
    }
}

/// Workbook read configuration shared by XLSX, XLS, and CSV engines.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReadOptions {
    /// Sheet selection.
    pub sheet: SheetSelector,
    /// Number of header rows. The final header row is used for name mapping.
    pub head_row_number: u32,
    /// Whether rows containing only empty cells are ignored.
    pub ignore_empty_row: bool,
    /// Whether leading and trailing whitespace is removed from string cells.
    pub auto_trim: bool,
    /// Whether numeric dates use Excel's 1904 windowing system.
    pub use_1904_windowing: bool,
    /// General-format rendering mode for extreme numbers.
    pub scientific_format: ScientificFormatMode,
    /// Locale used for formatted numeric and date display values.
    pub locale: ExcelLocale,
    /// Physical first row dispatched as data, zero-based and inclusive.
    ///
    /// Header rows are still analysed so name-based mapping remains available.
    pub start_row: Option<u32>,
    /// Physical last row dispatched as data, zero-based and inclusive.
    ///
    /// Header rows are still analysed so name-based mapping remains available.
    pub end_row: Option<u32>,
    /// Header aliases applied after optional Java-compatible trimming.
    ///
    /// Keys are workbook header names and values are names exposed to row mapping
    /// and `ReadListener::invoke_head`.
    pub header_aliases: HashMap<String, String>,
    /// User value exposed through every [`AnalysisContext`].
    pub custom_object: Option<CustomReadObject>,
    /// Value mode used by Java-compatible no-model [`easyexcel_core::DynamicRow`] reads.
    pub read_default_return: ReadDefaultReturn,
    /// Extra worksheet metadata dispatched to `ReadListener::extra`.
    pub extra_read: HashSet<CellExtraType>,
    /// Password used to decrypt an encrypted OOXML workbook.
    pub password: Option<String>,
    /// Character encoding used when reading CSV input.
    pub charset: CsvCharset,
    /// Java-style globally registered converters.
    pub converters: ConverterRegistry,
}

impl Default for ReadOptions {
    fn default() -> Self {
        Self {
            sheet: SheetSelector::First,
            head_row_number: 1,
            ignore_empty_row: true,
            auto_trim: true,
            use_1904_windowing: false,
            scientific_format: ScientificFormatMode::Plain,
            locale: ExcelLocale::default(),
            start_row: None,
            end_row: None,
            header_aliases: HashMap::new(),
            custom_object: None,
            read_default_return: ReadDefaultReturn::default(),
            extra_read: HashSet::new(),
            password: None,
            charset: CsvCharset::default(),
            converters: ConverterRegistry::default(),
        }
    }
}

/// Reads selected XLSX sheets and dispatches typed row events.
///
/// # Errors
///
/// Returns an I/O, workbook-format, sheet-selection, conversion, or listener error.
pub fn read_xlsx<T, L>(path: &Path, options: &ReadOptions, listener: &mut L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let source = open_xlsx_source(path, options)?;
    read_xlsx_source::<T, L>(&source, options, listener)
}

fn open_xlsx_source(path: &Path, options: &ReadOptions) -> Result<XlsxSource> {
    validate_read_options(options)?;
    XlsxSource::open(path, options.password.as_deref())
}

fn read_xlsx_source<T, L>(
    source: &XlsxSource,
    options: &ReadOptions,
    listener: &mut L,
) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let mut consumer = TypedRowConsumer::<T> { listener };
    read_xlsx_source_with_consumer(source, options, T::schema().is_empty(), &mut consumer)
}

fn read_xlsx_source_with_consumer(
    source: &XlsxSource,
    options: &ReadOptions,
    needs_cell_metadata: bool,
    consumer: &mut dyn RowConsumer,
) -> Result<()> {
    let mut workbook = Xlsx::new(source.reader()?).map_err(format_error)?;
    let mut row_metadata =
        (!options.ignore_empty_row || !options.extra_read.is_empty() || needs_cell_metadata)
            .then(|| source.reader().and_then(XlsxRowMetadata::new))
            .transpose()?;
    let names = selected_sheet_names(&workbook, &options.sheet, options.auto_trim)?;
    for (sheet_no, sheet_name) in names {
        let (last_explicit_row, extras) = xlsx_sheet_metadata(
            row_metadata.as_mut(),
            &sheet_name,
            options.ignore_empty_row,
            &options.extra_read,
        )?;
        let mut display_reader = if needs_cell_metadata {
            Some(
                row_metadata
                    .as_mut()
                    .expect("display metadata was initialized")
                    .display_cells(
                        &sheet_name,
                        options.use_1904_windowing,
                        options.scientific_format.is_enabled(),
                        options.locale.formatter(),
                    )?,
            )
        } else {
            None
        };
        if read_sheet(
            &mut workbook,
            sheet_no,
            &sheet_name,
            last_explicit_row,
            &extras,
            display_reader.as_mut(),
            options,
            consumer,
        )? == ReadFlow::Stop
        {
            break;
        }
    }
    Ok(())
}

fn xlsx_sheet_metadata(
    metadata: Option<&mut XlsxRowMetadata>,
    sheet_name: &str,
    ignore_empty_row: bool,
    enabled_extras: &HashSet<CellExtraType>,
) -> Result<(Option<u32>, Vec<CellExtra>)> {
    let Some(metadata) = metadata else {
        return Ok((None, Vec::new()));
    };
    let last_explicit_row = if ignore_empty_row {
        None
    } else {
        metadata.last_explicit_row(sheet_name)?
    };
    let extras = metadata.extras(sheet_name, enabled_extras)?;
    Ok((last_explicit_row, extras))
}

enum XlsxInput {
    File(BufReader<File>),
    Memory(Cursor<Arc<[u8]>>),
}

impl Read for XlsxInput {
    fn read(&mut self, buffer: &mut [u8]) -> std::io::Result<usize> {
        match self {
            Self::File(reader) => reader.read(buffer),
            Self::Memory(reader) => reader.read(buffer),
        }
    }
}

impl Seek for XlsxInput {
    fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
        match self {
            Self::File(reader) => reader.seek(position),
            Self::Memory(reader) => reader.seek(position),
        }
    }
}

enum XlsxSource {
    File(std::path::PathBuf),
    Memory(Arc<[u8]>),
}

impl XlsxSource {
    fn open(path: &Path, password: Option<&str>) -> Result<Self> {
        let mut reader = BufReader::new(File::open(path)?);
        // If the lightweight probe itself fails, the XLSX parser below still
        // returns the authoritative workbook error from the unchanged stream.
        if !is_compound_document(&mut reader) {
            return Ok(Self::File(path.to_owned()));
        }
        let password = password.ok_or_else(|| {
            ExcelError::Unsupported("encrypted OOXML workbook requires a password".to_owned())
        })?;
        let decrypted = match office_crypto::decrypt_from_file(path, password) {
            Ok(decrypted) => decrypted,
            Err(error) => return Err(decryption_error(&error)),
        };
        Ok(Self::Memory(Arc::from(decrypted)))
    }

    fn reader(&self) -> Result<XlsxInput> {
        match self {
            Self::File(path) => Ok(XlsxInput::File(BufReader::new(File::open(path)?))),
            Self::Memory(bytes) => Ok(XlsxInput::Memory(Cursor::new(Arc::clone(bytes)))),
        }
    }
}

fn is_compound_document(reader: &mut dyn BufRead) -> bool {
    const MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    reader
        .fill_buf()
        .is_ok_and(|actual| actual.len() >= MAGIC.len() && actual[..MAGIC.len()] == MAGIC)
}

fn decryption_error(error: &office_crypto::DecryptError) -> ExcelError {
    ExcelError::Format(format!("cannot decrypt workbook: {error}"))
}

/// Reads selected legacy XLS sheets through the typed listener lifecycle.
///
/// Calamine materializes each XLS worksheet before row dispatch because the
/// binary BIFF format does not expose the XLSX cell-stream API.
///
/// # Errors
///
/// Returns an I/O, workbook-format, sheet-selection, conversion, or listener error.
pub fn read_xls<T, L>(path: &Path, options: &ReadOptions, listener: &mut L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    reject_extra_read(options, "XLS")?;
    let mut workbook: Xls<_> = open_workbook(path).map_err(format_error)?;
    let sheets = select_xls_sheets(workbook.worksheets(), &options.sheet, options.auto_trim)?;
    for (sheet_no, sheet_name, range) in sheets {
        let mut consumer = TypedRowConsumer::<T> { listener };
        if read_range(&range, sheet_no, &sheet_name, options, &mut consumer)? == ReadFlow::Stop {
            break;
        }
    }
    Ok(())
}

/// Reads a CSV file through the same typed listener lifecycle as XLSX.
///
/// CSV exposes one logical sheet. Indexes other than zero return `SheetNotFound`.
///
/// # Errors
///
/// Returns an I/O, CSV-format, sheet-selection, conversion, or listener error.
pub fn read_csv<T, L>(path: &Path, options: &ReadOptions, listener: &mut L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    reject_extra_read(options, "CSV")?;
    let sheet_name = csv_sheet_name(&options.sheet)?;
    let encoding = csv_encoding(&options.charset)?;
    let input = DecodeReaderBytesBuilder::new()
        .encoding(Some(encoding))
        .strip_bom(true)
        .build(File::open(path)?);
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(input);
    read_csv_records::<T, L>(&mut reader.records(), 0, &sheet_name, options, listener)
}

fn csv_encoding(charset: &CsvCharset) -> Result<&'static Encoding> {
    Encoding::for_label(charset.name().as_bytes()).ok_or_else(|| {
        ExcelError::Unsupported(format!("unsupported CSV charset: {}", charset.name()))
    })
}

fn read_csv_records<T, L>(
    records: &mut dyn Iterator<Item = csv::Result<csv::StringRecord>>,
    start_row: usize,
    sheet_name: &str,
    options: &ReadOptions,
    listener: &mut L,
) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let mut headers = Arc::new(HashMap::new());
    let mut final_row = 0_u32;
    for (offset, record) in records.enumerate() {
        let row_index = start_row.saturating_add(offset);
        let row_index = csv_row_index(row_index)?;
        final_row = row_index;
        let cells = record
            .map_err(format_error)?
            .iter()
            .map(|value| CellValue::String(value.to_owned()))
            .collect();
        if process_row::<T>(
            0,
            sheet_name,
            row_index,
            cells,
            options,
            &mut headers,
            listener,
        )? == ReadFlow::Stop
        {
            return Ok(());
        }
    }
    listener.do_after_all_analysed(&analysis_context(sheet_name, 0, final_row, options))
}

fn csv_row_index(row_index: usize) -> Result<u32> {
    u32::try_from(row_index).map_err(|_| ExcelError::Format("CSV row index exceeds u32".to_owned()))
}

fn csv_sheet_name(selector: &SheetSelector) -> Result<String> {
    match selector {
        SheetSelector::First | SheetSelector::Index(0) | SheetSelector::All => {
            Ok("Sheet1".to_owned())
        }
        SheetSelector::Name(name) => Ok(name.clone()),
        SheetSelector::Index(index) => Err(ExcelError::SheetNotFound(index.to_string())),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReadFlow {
    Continue,
    Stop,
}

#[derive(Default)]
struct SourceRowMetadata {
    formulas: HashMap<usize, FormulaData>,
    display_values: HashMap<usize, String>,
    decimal_values: HashMap<usize, bigdecimal::BigDecimal>,
    present_columns: HashSet<usize>,
}

trait RowConsumer {
    #[allow(clippy::too_many_arguments)]
    fn process(
        &mut self,
        sheet_no: usize,
        sheet_name: &str,
        row_index: u32,
        cells: Vec<CellValue>,
        metadata: SourceRowMetadata,
        options: &ReadOptions,
        headers: &mut Arc<HashMap<String, usize>>,
    ) -> Result<ReadFlow>;

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<ReadFlow>;

    fn after(&mut self, context: &AnalysisContext) -> Result<()>;
}

struct TypedRowConsumer<'a, T> {
    listener: &'a mut dyn ReadListener<T>,
}

impl<T: ExcelRow> RowConsumer for TypedRowConsumer<'_, T> {
    fn process(
        &mut self,
        sheet_no: usize,
        sheet_name: &str,
        row_index: u32,
        cells: Vec<CellValue>,
        metadata: SourceRowMetadata,
        options: &ReadOptions,
        headers: &mut Arc<HashMap<String, usize>>,
    ) -> Result<ReadFlow> {
        process_row_with_metadata::<T>(
            sheet_no,
            sheet_name,
            row_index,
            cells,
            metadata,
            options,
            headers,
            self.listener,
        )
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<ReadFlow> {
        let result = self.listener.extra(extra, context);
        listener_result(result, self.listener, context)
    }

    fn after(&mut self, context: &AnalysisContext) -> Result<()> {
        self.listener.do_after_all_analysed(context)
    }
}

fn selected_sheet_names<RS: std::io::Read + std::io::Seek>(
    workbook: &Xlsx<RS>,
    selector: &SheetSelector,
    auto_trim: bool,
) -> Result<Vec<(usize, String)>> {
    select_sheet_names(workbook.sheet_names(), selector, auto_trim)
}

fn select_sheet_names(
    names: Vec<String>,
    selector: &SheetSelector,
    auto_trim: bool,
) -> Result<Vec<(usize, String)>> {
    match selector {
        SheetSelector::First => names
            .first()
            .cloned()
            .map(|name| vec![(0, name)])
            .ok_or_else(|| ExcelError::SheetNotFound("0".to_owned())),
        SheetSelector::Index(index) => names
            .get(*index)
            .cloned()
            .map(|name| vec![(*index, name)])
            .ok_or_else(|| ExcelError::SheetNotFound(index.to_string())),
        SheetSelector::Name(name) => names
            .iter()
            .enumerate()
            .find(|(_, candidate)| sheet_name_matches(candidate, name, auto_trim))
            .map(|(index, candidate)| vec![(index, candidate.clone())])
            .ok_or_else(|| ExcelError::SheetNotFound(name.clone())),
        SheetSelector::All => Ok(names.into_iter().enumerate().collect()),
    }
}

fn select_xls_sheets(
    sheets: Vec<(String, Range<Data>)>,
    selector: &SheetSelector,
    auto_trim: bool,
) -> Result<Vec<(usize, String, Range<Data>)>> {
    match selector {
        SheetSelector::First => sheets
            .into_iter()
            .next()
            .map(|(name, range)| vec![(0, name, range)])
            .ok_or_else(|| ExcelError::SheetNotFound("0".to_owned())),
        SheetSelector::Index(index) => sheets
            .into_iter()
            .nth(*index)
            .map(|(name, range)| vec![(*index, name, range)])
            .ok_or_else(|| ExcelError::SheetNotFound(index.to_string())),
        SheetSelector::Name(name) => sheets
            .into_iter()
            .enumerate()
            .find(|(_, (candidate, _))| sheet_name_matches(candidate, name, auto_trim))
            .map(|(index, (candidate, range))| vec![(index, candidate, range)])
            .ok_or_else(|| ExcelError::SheetNotFound(name.clone())),
        SheetSelector::All => Ok(sheets
            .into_iter()
            .enumerate()
            .map(|(index, (name, range))| (index, name, range))
            .collect()),
    }
}

#[allow(clippy::too_many_arguments, clippy::too_many_lines)]
fn read_sheet<RS>(
    workbook: &mut Xlsx<RS>,
    sheet_no: usize,
    sheet_name: &str,
    last_explicit_row: Option<u32>,
    extras: &[CellExtra],
    mut display_reader: Option<&mut XlsxDisplayCellReader<'_>>,
    options: &ReadOptions,
    consumer: &mut dyn RowConsumer,
) -> Result<ReadFlow>
where
    RS: std::io::Read + std::io::Seek,
{
    let mut reader = workbook
        .worksheet_cells_reader(sheet_name)
        .map_err(format_error)?;
    let mut current_index = None;
    let mut current_cells = Vec::new();
    let mut current_formulas = HashMap::new();
    let mut current_display_values = HashMap::new();
    let mut current_decimal_values = HashMap::new();
    let mut current_present_columns = HashSet::new();
    let mut headers = Arc::new(HashMap::new());
    let mut next_row_index = 0;

    while let Some(cell) = reader.next_cell_with_formula().map_err(format_error)? {
        let (row, column) = cell.pos;
        if current_index != Some(row) {
            if let Some(current) = current_index {
                if dispatch_row(
                    consumer,
                    sheet_no,
                    sheet_name,
                    current,
                    std::mem::take(&mut current_cells),
                    SourceRowMetadata {
                        formulas: std::mem::take(&mut current_formulas),
                        display_values: std::mem::take(&mut current_display_values),
                        decimal_values: std::mem::take(&mut current_decimal_values),
                        present_columns: std::mem::take(&mut current_present_columns),
                    },
                    options,
                    &mut headers,
                )? == ReadFlow::Stop
                {
                    return Ok(ReadFlow::Stop);
                }
                next_row_index = current.saturating_add(1);
            }
            if process_missing_rows(
                next_row_index,
                row,
                sheet_no,
                sheet_name,
                options,
                &mut headers,
                consumer,
            )? == ReadFlow::Stop
            {
                return Ok(ReadFlow::Stop);
            }
            current_index = Some(row);
        }
        let column = to_column_index(column)?;
        if let Some(display_reader) = display_reader.as_deref_mut() {
            let display_cell = display_reader.next_cell()?.ok_or_else(|| {
                ExcelError::Format("display-value stream ended before cell stream".to_owned())
            })?;
            if display_cell.position != (row, column) {
                return Err(ExcelError::Format(format!(
                    "display-value stream cell mismatch: expected ({row}, {column}), found ({}, {})",
                    display_cell.position.0, display_cell.position.1
                )));
            }
            if let Some(value) = display_cell.display_value {
                current_display_values.insert(column, value);
            }
            if let Some(value) = display_cell.decimal_value {
                current_decimal_values.insert(column, value);
            }
        }
        if current_cells.len() <= column {
            current_cells.resize(column + 1, CellValue::Empty);
        }
        current_present_columns.insert(column);
        current_cells[column] = from_calamine(&cell.value, options.use_1904_windowing);
        if let Some(formula) = cell.formula {
            current_formulas.insert(column, FormulaData::new(formula));
        }
    }

    if let Some(row) = current_index
        && dispatch_row(
            consumer,
            sheet_no,
            sheet_name,
            row,
            current_cells,
            SourceRowMetadata {
                formulas: current_formulas,
                display_values: current_display_values,
                decimal_values: current_decimal_values,
                present_columns: current_present_columns,
            },
            options,
            &mut headers,
        )? == ReadFlow::Stop
    {
        return Ok(ReadFlow::Stop);
    }

    if let Some(display_reader) = display_reader
        && display_reader.next_cell()?.is_some()
    {
        return Err(ExcelError::Format(
            "display-value stream contains cells missing from cell stream".to_owned(),
        ));
    }

    if let Some(last_row) = last_explicit_row {
        let first_trailing_row = current_index.map_or(0, |row| row.saturating_add(1));
        if process_missing_rows(
            first_trailing_row,
            last_row.saturating_add(1),
            sheet_no,
            sheet_name,
            options,
            &mut headers,
            consumer,
        )? == ReadFlow::Stop
        {
            return Ok(ReadFlow::Stop);
        }
    }

    let final_row = last_explicit_row.or(current_index).unwrap_or_default();
    let context = analysis_context(sheet_name, sheet_no, final_row, options);
    for extra in extras {
        if consumer.extra(extra, &context)? == ReadFlow::Stop {
            return Ok(ReadFlow::Stop);
        }
    }
    consumer.after(&context)?;
    Ok(ReadFlow::Continue)
}

#[allow(clippy::too_many_arguments)]
fn process_missing_rows(
    start_row: u32,
    end_row: u32,
    sheet_no: usize,
    sheet_name: &str,
    options: &ReadOptions,
    headers: &mut Arc<HashMap<String, usize>>,
    consumer: &mut dyn RowConsumer,
) -> Result<ReadFlow> {
    for row_index in start_row..end_row {
        if dispatch_row(
            consumer,
            sheet_no,
            sheet_name,
            row_index,
            Vec::new(),
            SourceRowMetadata::default(),
            options,
            headers,
        )? == ReadFlow::Stop
        {
            return Ok(ReadFlow::Stop);
        }
    }
    Ok(ReadFlow::Continue)
}

fn read_range(
    range: &Range<Data>,
    sheet_no: usize,
    sheet_name: &str,
    options: &ReadOptions,
    consumer: &mut dyn RowConsumer,
) -> Result<ReadFlow> {
    let mut headers = Arc::new(HashMap::new());
    let Some((start_row, start_column)) = range.start() else {
        consumer.after(&analysis_context(sheet_name, sheet_no, 0, options))?;
        return Ok(ReadFlow::Continue);
    };
    let start_column = to_column_index(start_column)?;
    let mut row_index = start_row;
    let mut final_row = start_row;
    for row in range.rows() {
        final_row = row_index;
        let mut cells = vec![CellValue::Empty; start_column];
        cells.extend(
            row.iter()
                .map(|value| from_data(value, options.use_1904_windowing)),
        );
        let present_columns = row
            .iter()
            .enumerate()
            .filter_map(|(offset, value)| {
                (!matches!(value, Data::Empty)).then_some(start_column + offset)
            })
            .collect();
        if dispatch_row(
            consumer,
            sheet_no,
            sheet_name,
            row_index,
            cells,
            SourceRowMetadata {
                present_columns,
                ..SourceRowMetadata::default()
            },
            options,
            &mut headers,
        )? == ReadFlow::Stop
        {
            return Ok(ReadFlow::Stop);
        }
        row_index = row_index.saturating_add(1);
    }
    consumer.after(&analysis_context(sheet_name, sheet_no, final_row, options))?;
    Ok(ReadFlow::Continue)
}

#[allow(clippy::too_many_arguments)]
fn process_row<T>(
    sheet_no: usize,
    sheet_name: &str,
    row_index: u32,
    cells: Vec<CellValue>,
    options: &ReadOptions,
    headers: &mut Arc<HashMap<String, usize>>,
    listener: &mut dyn ReadListener<T>,
) -> Result<ReadFlow>
where
    T: ExcelRow,
{
    let present_columns = (0..cells.len()).collect();
    let mut consumer = TypedRowConsumer::<T> { listener };
    dispatch_row(
        &mut consumer,
        sheet_no,
        sheet_name,
        row_index,
        cells,
        SourceRowMetadata {
            present_columns,
            ..SourceRowMetadata::default()
        },
        options,
        headers,
    )
}

#[allow(clippy::too_many_arguments)]
fn dispatch_row(
    consumer: &mut dyn RowConsumer,
    sheet_no: usize,
    sheet_name: &str,
    row_index: u32,
    cells: Vec<CellValue>,
    metadata: SourceRowMetadata,
    options: &ReadOptions,
    headers: &mut Arc<HashMap<String, usize>>,
) -> Result<ReadFlow> {
    if row_index >= options.head_row_number
        && (options.start_row.is_some_and(|start| row_index < start)
            || options.end_row.is_some_and(|end| row_index > end))
    {
        return Ok(ReadFlow::Continue);
    }
    consumer.process(
        sheet_no, sheet_name, row_index, cells, metadata, options, headers,
    )
}

fn analysis_context(
    sheet_name: &str,
    sheet_no: usize,
    row_index: u32,
    options: &ReadOptions,
) -> AnalysisContext {
    AnalysisContext::new(sheet_name, sheet_no, row_index)
        .with_custom_object(options.custom_object.clone())
}

#[allow(clippy::too_many_arguments)]
fn process_row_with_metadata<T>(
    sheet_no: usize,
    sheet_name: &str,
    row_index: u32,
    mut cells: Vec<CellValue>,
    metadata: SourceRowMetadata,
    options: &ReadOptions,
    headers: &mut Arc<HashMap<String, usize>>,
    listener: &mut dyn ReadListener<T>,
) -> Result<ReadFlow>
where
    T: ExcelRow,
{
    let SourceRowMetadata {
        formulas,
        display_values,
        decimal_values,
        present_columns,
    } = metadata;
    if options.auto_trim {
        trim_string_cells(&mut cells);
    }
    let context = analysis_context(sheet_name, sheet_no, row_index, options);
    if row_index < options.head_row_number {
        let current_headers = Arc::new(header_map(&cells, &options.header_aliases));
        if row_index + 1 == options.head_row_number {
            *headers = Arc::clone(&current_headers);
        }
        let result = listener.invoke_head(&current_headers, &context);
        return listener_result(result, listener, &context);
    }
    if options.ignore_empty_row && cells.iter().all(is_empty_read_cell) {
        return Ok(ReadFlow::Continue);
    }

    let row = RowData::new(sheet_name, row_index, cells, Arc::clone(headers))
        .with_formulas(formulas)
        .with_display_values(display_values)
        .with_decimal_values(decimal_values)
        .with_present_columns(present_columns)
        .with_read_default_return(options.read_default_return);
    match T::from_row_with_converters(&row, &options.converters) {
        Ok(data) => {
            let result = listener.invoke(data, &context);
            listener_result(result, listener, &context)
        }
        Err(error) => listener_error(error, listener, &context),
    }
}

fn trim_string_cells(cells: &mut [CellValue]) {
    for cell in cells {
        if let CellValue::String(value) = cell {
            let trimmed = java_trim(value);
            if trimmed.len() != value.len() {
                *value = trimmed.to_owned();
            }
        }
    }
}

fn is_empty_read_cell(cell: &CellValue) -> bool {
    cell.is_empty() || matches!(cell, CellValue::String(value) if value.is_empty())
}

fn sheet_name_matches(candidate: &str, requested: &str, auto_trim: bool) -> bool {
    if auto_trim {
        java_trim(candidate) == java_trim(requested)
    } else {
        candidate == requested
    }
}

fn reject_extra_read(options: &ReadOptions, format: &str) -> Result<()> {
    validate_read_options(options)?;
    if options.extra_read.is_empty() {
        Ok(())
    } else {
        Err(ExcelError::Unsupported(format!(
            "{format} extra metadata is not supported"
        )))
    }
}

fn java_trim(value: &str) -> &str {
    value.trim_matches(|character| character <= '\u{20}')
}

fn listener_result<T>(
    result: Result<()>,
    listener: &mut dyn ReadListener<T>,
    context: &AnalysisContext,
) -> Result<ReadFlow> {
    match result {
        Ok(()) if listener.has_next(context) => Ok(ReadFlow::Continue),
        Ok(()) => Ok(ReadFlow::Stop),
        Err(error) => listener_error(error, listener, context),
    }
}

fn listener_error<T>(
    error: ExcelError,
    listener: &mut dyn ReadListener<T>,
    context: &AnalysisContext,
) -> Result<ReadFlow> {
    match listener.on_exception(&error, context) {
        ErrorAction::Continue | ErrorAction::SkipRow => Ok(ReadFlow::Continue),
        ErrorAction::Stop => Err(error),
    }
}

fn validate_read_options(options: &ReadOptions) -> Result<()> {
    if let (Some(start), Some(end)) = (options.start_row, options.end_row)
        && start > end
    {
        return Err(ExcelError::Format(format!(
            "read row range start {start} exceeds end {end}"
        )));
    }
    Ok(())
}

fn header_map(
    cells: &[CellValue],
    header_aliases: &HashMap<String, String>,
) -> HashMap<String, usize> {
    cells
        .iter()
        .enumerate()
        .filter_map(|(index, value)| {
            let name = value.as_text();
            (!name.is_empty()).then(|| {
                let alias = header_aliases.get(&name).cloned().unwrap_or(name);
                (alias, index)
            })
        })
        .collect()
}

fn from_calamine(value: &DataRef<'_>, use_1904_windowing: bool) -> CellValue {
    match value {
        DataRef::Empty => CellValue::Empty,
        DataRef::String(value) | DataRef::DateTimeIso(value) | DataRef::DurationIso(value) => {
            CellValue::String(value.clone())
        }
        DataRef::SharedString(value) => CellValue::String((*value).to_owned()),
        DataRef::Bool(value) => CellValue::Bool(*value),
        DataRef::Int(value) => CellValue::Int(*value),
        DataRef::Float(value) => CellValue::Float(*value),
        DataRef::DateTime(value) => excel_datetime_cell(value, use_1904_windowing),
        DataRef::Error(value) => CellValue::String(value.to_string()),
    }
}

fn from_data(value: &Data, use_1904_windowing: bool) -> CellValue {
    match value {
        Data::Empty => CellValue::Empty,
        Data::String(value) | Data::DateTimeIso(value) | Data::DurationIso(value) => {
            CellValue::String(value.clone())
        }
        Data::Bool(value) => CellValue::Bool(*value),
        Data::Int(value) => CellValue::Int(*value),
        Data::Float(value) => CellValue::Float(*value),
        Data::DateTime(value) => excel_datetime_cell(value, use_1904_windowing),
        Data::Error(value) => CellValue::String(value.to_string()),
    }
}

fn excel_datetime_cell(value: &ExcelDateTime, use_1904_windowing: bool) -> CellValue {
    if !value.is_datetime() {
        return CellValue::Float(value.as_f64());
    }
    ExcelDateTime::new(
        value.as_f64(),
        ExcelDateTimeType::DateTime,
        use_1904_windowing,
    )
    .as_datetime()
    .map_or(CellValue::Float(value.as_f64()), CellValue::DateTime)
}

fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

fn to_column_index(column: u32) -> Result<usize> {
    u16::try_from(column)
        .map(usize::from)
        .map_err(|_| ExcelError::Format("column index exceeds spreadsheet limit".to_owned()))
}

#[cfg(test)]
mod tests;
