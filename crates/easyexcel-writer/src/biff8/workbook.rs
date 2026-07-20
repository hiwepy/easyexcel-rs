//! In-memory BIFF8 workbook model and OLE/CFB serialization.
//!
//! Java mapping: Alibaba EasyExcel `excelType(ExcelTypeEnum.XLS)` → POI HSSF.
//! This module is a **minimal** BIFF8 writer (not a full HSSF port):
//! - Supported: single/multi sheet, header + data rows, string / number / bool /
//!   date / datetime cells, SST shared strings, 1900 date system, column widths
//!   (COLINFO), row heights (ROW), basic FONT/XF (bold/italic/size/indexed or
//!   approximated RGB fill), MERGECELLS ranges.
//! - Template / scalar fill: value-preserving rewrite via [`super::template`]
//!   (styles/merges not preserved). Collection fill and in-place OLE patching remain
//!   unsupported. Also unsupported: password encryption, images, true formula tokens,
//!   hyperlink/comment records, rich-text runs, borders, arbitrary custom number
//!   formats, charts, macros. Gaps fail visibly — never silently rewrite as XLSX.

use std::collections::{BTreeMap, HashMap};
use std::io::{Cursor, Write};

use chrono::{NaiveDate, NaiveDateTime, Timelike};
use easyexcel_core::{ExcelError, Result};

use super::encode::{
    encode_rk, encode_short_unicode_string, encode_unicode_string, pack_colinfo, pack_merge_range,
    pack_row, record, write_merge_cells, write_palette_record, BIFF8_VERSION, BLANK, BOOLERR,
    BOUNDSHEET, BOF, CODEPAGE, COLINFO, CONTINUE, DATEMODE, DIMENSION, DT_GLOBALS, DT_WORKSHEET,
    EOF, EXTSST, FONT, LABELSST, MAX_RECORD_DATA, NUMBER, RK, ROW, SST, STYLE, WINDOW2, XF, XF_DATE,
    XF_DATETIME, XF_GENERAL,
};
use super::style::Biff8StyleTable;

/// A cell value ready for BIFF8 emission, with an XF index for date formats.
#[derive(Debug, Clone)]
pub struct Biff8Cell {
    /// Logical value.
    pub value: Biff8Value,
    /// XF index (`XF_GENERAL` / `XF_DATE` / `XF_DATETIME` / custom ≥ 18).
    pub xf: u16,
}

impl Biff8Cell {
    /// Creates a general-format cell.
    #[must_use]
    pub const fn general(value: Biff8Value) -> Self {
        Self {
            value,
            xf: XF_GENERAL,
        }
    }

    /// Creates a date-formatted numeric serial cell.
    #[must_use]
    pub const fn date_serial(serial: f64) -> Self {
        Self {
            value: Biff8Value::Number(serial),
            xf: XF_DATE,
        }
    }

    /// Creates a datetime-formatted numeric serial cell.
    #[must_use]
    pub const fn datetime_serial(serial: f64) -> Self {
        Self {
            value: Biff8Value::Number(serial),
            xf: XF_DATETIME,
        }
    }

    /// Returns a copy with a different XF index (styled date/general cells).
    #[must_use]
    pub const fn with_xf(mut self, xf: u16) -> Self {
        self.xf = xf;
        self
    }
}

/// Logical cell payload (before SST / NUMBER framing).
#[derive(Debug, Clone)]
pub enum Biff8Value {
    /// Blank cell.
    Blank,
    /// Shared string (interned into SST on serialize).
    Text(String),
    /// IEEE754 number (also used for Excel date serials).
    Number(f64),
    /// Boolean.
    Bool(bool),
}

/// One inclusive merge region in BIFF coordinates (Java HSSF `CellRangeAddress`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Biff8Merge {
    /// First row (0-based).
    pub first_row: u16,
    /// Last row (0-based, inclusive).
    pub last_row: u16,
    /// First column (0-based).
    pub first_col: u8,
    /// Last column (0-based, inclusive).
    pub last_col: u8,
}

/// Sparse worksheet buffer accumulated by the stateful / one-shot writers.
#[derive(Debug, Clone, Default)]
pub struct Biff8Sheet {
    /// Worksheet name (BOUNDSHEET short Unicode string).
    pub name: String,
    /// Sparse cells keyed by `(row, column)` in 0-based BIFF coordinates.
    pub cells: BTreeMap<(u16, u8), Biff8Cell>,
    /// Column widths in Excel character units (Java `sheet.setColumnWidth`).
    pub column_widths: BTreeMap<u8, u16>,
    /// Row heights in points (Java `row.setHeightInPoints`).
    pub row_heights: BTreeMap<u16, u16>,
    /// Merged regions (Java `addMergedRegion` / `MergedCellsTable`).
    pub merges: Vec<Biff8Merge>,
    /// Next free row index (includes any header rows already written).
    pub next_row: u32,
    /// Next data-row index used for content-style cycling parity with XLSX.
    pub next_data_index: usize,
}

impl Biff8Sheet {
    /// Creates an empty sheet with the given name.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: BTreeMap::new(),
            column_widths: BTreeMap::new(),
            row_heights: BTreeMap::new(),
            merges: Vec::new(),
            next_row: 0,
            next_data_index: 0,
        }
    }

    /// Writes a cell at `(row, col)`, enforcing BIFF8 row/column limits.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Format`] when the coordinate exceeds BIFF8 limits
    /// (65_536 rows × 256 columns).
    pub fn set(&mut self, row: u32, col: usize, cell: Biff8Cell) -> Result<()> {
        let row = u16::try_from(row).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 65536 rows".to_owned())
        })?;
        let col = u8::try_from(col).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 256 columns".to_owned())
        })?;
        self.cells.insert((row, col), cell);
        Ok(())
    }

    /// Sets column width in character units (POI `setColumnWidth(col, chars*256)`).
    pub fn set_column_width(&mut self, col: u8, width_chars: u16) {
        self.column_widths.insert(col, width_chars);
    }

    /// Sets row height in points (POI `setHeightInPoints`).
    pub fn set_row_height(&mut self, row: u16, height_points: u16) {
        self.row_heights.insert(row, height_points);
    }

    /// Appends a merge region when it spans more than one cell.
    pub fn add_merge(&mut self, merge: Biff8Merge) -> Result<()> {
        if merge.last_row < merge.first_row || merge.last_col < merge.first_col {
            return Err(ExcelError::Format(
                "BIFF8 merge last row/col must be >= first".to_owned(),
            ));
        }
        if merge.first_row == merge.last_row && merge.first_col == merge.last_col {
            return Ok(());
        }
        self.merges.push(merge);
        Ok(())
    }

    /// Returns exclusive `(max_row, max_col)` for the DIMENSION record.
    fn dimensions(&self) -> (u32, u16) {
        let mut max_row = 0u32;
        let mut max_col = 0u16;
        for &(row, col) in self.cells.keys() {
            max_row = max_row.max(u32::from(row).saturating_add(1));
            max_col = max_col.max(u16::from(col).saturating_add(1));
        }
        for merge in &self.merges {
            max_row = max_row.max(u32::from(merge.last_row).saturating_add(1));
            max_col = max_col.max(u16::from(merge.last_col).saturating_add(1));
        }
        (max_row, max_col)
    }
}

/// Multi-sheet BIFF8 workbook buffer.
#[derive(Debug, Clone, Default)]
pub struct Biff8Book {
    /// Ordered worksheets (emission order = BOUNDSHEET order).
    pub sheets: Vec<Biff8Sheet>,
    /// Workbook-global FONT / XF registry (Java HSSF style table).
    pub styles: Biff8StyleTable,
    /// When `true`, BIFF8 `DATEMODE` uses the 1904 date windowing system.
    pub use_1904_windowing: bool,
}

impl Biff8Book {
    /// Returns a mutable sheet by name, creating it if missing.
    pub fn sheet_mut(&mut self, name: &str) -> &mut Biff8Sheet {
        if let Some(index) = self.sheets.iter().position(|s| s.name == name) {
            return &mut self.sheets[index];
        }
        self.sheets.push(Biff8Sheet::new(name.to_owned()));
        self.sheets.last_mut().expect("just pushed")
    }

    /// Serializes this book to an OLE Compound File containing a `Workbook` stream.
    ///
    /// # Errors
    ///
    /// Returns I/O or CFB construction errors.
    pub fn to_cfb_bytes(&self) -> Result<Vec<u8>> {
        let stream = build_workbook_stream(self)?;
        let mut mem = Cursor::new(Vec::<u8>::new());
        {
            let mut cf = cfb::CompoundFile::create(&mut mem).map_err(|error| {
                ExcelError::Format(format!("cannot create OLE2 container: {error}"))
            })?;
            {
                let mut workbook = cf.create_stream("Workbook").map_err(|error| {
                    ExcelError::Format(format!("cannot create Workbook stream: {error}"))
                })?;
                workbook.write_all(&stream)?;
            }
            cf.flush().map_err(|error| {
                ExcelError::Format(format!("cannot flush OLE2 container: {error}"))
            })?;
        }
        Ok(mem.into_inner())
    }

    /// Writes the CFB bytes to `writer`.
    ///
    /// # Errors
    ///
    /// Returns serialization or I/O errors.
    pub fn write_to<W: Write>(&self, mut writer: W) -> Result<()> {
        let bytes = self.to_cfb_bytes()?;
        writer.write_all(&bytes)?;
        Ok(())
    }
}

/// Converts a calendar date to an Excel 1900-date-system serial number.
#[must_use]
pub fn date_to_excel_serial(date: NaiveDate) -> f64 {
    date_to_excel_serial_with_windowing(date, false)
}

/// Converts a calendar date using either the 1900 or 1904 date windowing system.
#[must_use]
pub fn date_to_excel_serial_with_windowing(date: NaiveDate, use_1904_windowing: bool) -> f64 {
    let epoch = if use_1904_windowing {
        // Excel 1904 system: day 0 is 1904-01-01.
        NaiveDate::from_ymd_opt(1904, 1, 1).expect("valid epoch")
    } else {
        // Excel's 1900 system epoch is 1899-12-30 (Lotus 1-2-3 leap-year bug).
        NaiveDate::from_ymd_opt(1899, 12, 30).expect("valid epoch")
    };
    f64::from(i32::try_from((date - epoch).num_days()).unwrap_or(i32::MAX))
}

/// Converts a naive date-time to an Excel serial (date + fraction of day).
#[must_use]
pub fn datetime_to_excel_serial(value: NaiveDateTime) -> f64 {
    datetime_to_excel_serial_with_windowing(value, false)
}

/// Converts a naive date-time using either the 1900 or 1904 date windowing system.
#[must_use]
pub fn datetime_to_excel_serial_with_windowing(
    value: NaiveDateTime,
    use_1904_windowing: bool,
) -> f64 {
    let date_part = date_to_excel_serial_with_windowing(value.date(), use_1904_windowing);
    let time = value.time();
    let seconds = f64::from(time.num_seconds_from_midnight())
        + f64::from(time.nanosecond()) / 1_000_000_000.0;
    date_part + seconds / 86_400.0
}

/// Builds the BIFF8 `Workbook` stream (globals + worksheet substreams).
fn build_workbook_stream(book: &Biff8Book) -> Result<Vec<u8>> {
    let mut out: Vec<u8> = Vec::new();
    write_bof(&mut out, DT_GLOBALS);
    record(&mut out, CODEPAGE, &1200u16.to_le_bytes());
    let date_mode = u16::from(book.use_1904_windowing);
    record(&mut out, DATEMODE, &date_mode.to_le_bytes());

    for _ in 0..5 {
        write_default_font(&mut out);
    }
    for font in book.styles.custom_fonts() {
        record(&mut out, FONT, &font);
    }
    if book.styles.needs_palette() {
        write_palette_record(&mut out, book.styles.palette_overrides());
    }
    for _ in 0..16 {
        write_style_xf(&mut out);
    }
    write_cell_xf(&mut out, 14); // XF_DATE
    write_cell_xf(&mut out, 22); // XF_DATETIME
    for xf in book.styles.custom_xfs() {
        record(&mut out, XF, xf);
    }

    {
        let mut data = Vec::new();
        data.extend_from_slice(&0x8000u16.to_le_bytes());
        data.push(0x00);
        data.push(0xFF);
        record(&mut out, STYLE, &data);
    }

    let sheets = if book.sheets.is_empty() {
        vec![Biff8Sheet::new("Sheet1")]
    } else {
        book.sheets.clone()
    };
    let (sst_strings, sst_index, total_refs) = build_sst(&sheets);

    let mut boundsheet_patches = Vec::with_capacity(sheets.len());
    for sheet in &sheets {
        boundsheet_patches.push(write_boundsheet_placeholder(&mut out, &sheet.name));
    }

    if !sst_strings.is_empty() {
        out.extend_from_slice(&build_sst_records(&sst_strings, total_refs));
        record(&mut out, EXTSST, &[0, 0]);
    }
    record(&mut out, EOF, &[]);

    let mut sheet_offsets = Vec::with_capacity(sheets.len());
    for sheet in &sheets {
        sheet_offsets.push(out.len() as u32);
        write_worksheet(&mut out, sheet, &sst_index);
    }
    for (patch_off, pos) in boundsheet_patches.into_iter().zip(sheet_offsets) {
        out[patch_off..patch_off + 4].copy_from_slice(&pos.to_le_bytes());
    }
    Ok(out)
}

fn write_bof(out: &mut Vec<u8>, dt: u16) {
    let mut data = Vec::new();
    data.extend_from_slice(&BIFF8_VERSION.to_le_bytes());
    data.extend_from_slice(&dt.to_le_bytes());
    data.extend_from_slice(&0x0DBBu16.to_le_bytes());
    data.extend_from_slice(&0x07CCu16.to_le_bytes());
    data.extend_from_slice(&0x0000_00C1u32.to_le_bytes());
    data.extend_from_slice(&0x0000_0006u32.to_le_bytes());
    record(out, BOF, &data);
}

fn write_default_font(out: &mut Vec<u8>) {
    let mut data = Vec::new();
    data.extend_from_slice(&200u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&0x7FFFu16.to_le_bytes());
    data.extend_from_slice(&400u16.to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&[0, 0, 0, 0]);
    data.extend_from_slice(&encode_short_unicode_string("Arial"));
    record(out, FONT, &data);
}

fn write_style_xf(out: &mut Vec<u8>) {
    let mut data = vec![0u8; 20];
    data[4] = 0xF5;
    data[5] = 0xFF;
    record(out, XF, &data);
}

fn write_cell_xf(out: &mut Vec<u8>, ifmt: u16) {
    let mut data = vec![0u8; 20];
    data[2..4].copy_from_slice(&ifmt.to_le_bytes());
    data[4..6].copy_from_slice(&0x0001u16.to_le_bytes());
    record(out, XF, &data);
}

fn write_boundsheet_placeholder(out: &mut Vec<u8>, name: &str) -> usize {
    let mut data = Vec::new();
    data.extend_from_slice(&0u32.to_le_bytes());
    data.extend_from_slice(&[0x00, 0x00]);
    data.extend_from_slice(&encode_short_unicode_string(name));
    let record_start = out.len();
    record(out, BOUNDSHEET, &data);
    record_start + 4
}

fn build_sst(sheets: &[Biff8Sheet]) -> (Vec<String>, HashMap<String, u32>, u32) {
    let mut strings = Vec::new();
    let mut index = HashMap::new();
    let mut total_refs = 0u32;
    for sheet in sheets {
        for cell in sheet.cells.values() {
            if let Biff8Value::Text(text) = &cell.value {
                total_refs += 1;
                if let std::collections::hash_map::Entry::Vacant(entry) = index.entry(text.clone())
                {
                    entry.insert(strings.len() as u32);
                    strings.push(text.clone());
                }
            }
        }
    }
    (strings, index, total_refs)
}

fn build_sst_records(strings: &[String], total_refs: u32) -> Vec<u8> {
    let mut pieces: Vec<Vec<u8>> = Vec::new();
    let mut header = Vec::new();
    header.extend_from_slice(&total_refs.to_le_bytes());
    header.extend_from_slice(&(strings.len() as u32).to_le_bytes());
    pieces.push(header);
    for s in strings {
        pieces.push(encode_unicode_string(s));
    }

    let mut out = Vec::new();
    let mut current = Vec::new();
    let mut first = true;
    for piece in pieces {
        if !current.is_empty() && current.len() + piece.len() > MAX_RECORD_DATA {
            flush_sst_chunk(&mut out, &mut current, &mut first);
        }
        if piece.len() > MAX_RECORD_DATA {
            let mut offset = 0;
            while offset < piece.len() {
                let room = MAX_RECORD_DATA.saturating_sub(current.len());
                if room == 0 {
                    flush_sst_chunk(&mut out, &mut current, &mut first);
                    continue;
                }
                let take = room.min(piece.len() - offset);
                current.extend_from_slice(&piece[offset..offset + take]);
                offset += take;
            }
        } else {
            current.extend_from_slice(&piece);
        }
    }
    if !current.is_empty() {
        flush_sst_chunk(&mut out, &mut current, &mut first);
    }
    out
}

fn flush_sst_chunk(out: &mut Vec<u8>, current: &mut Vec<u8>, first: &mut bool) {
    let typ = if *first { SST } else { CONTINUE };
    record(out, typ, current);
    *first = false;
    current.clear();
}

fn write_worksheet(out: &mut Vec<u8>, sheet: &Biff8Sheet, sst_index: &HashMap<String, u32>) {
    write_bof(out, DT_WORKSHEET);
    let (max_row, max_col) = sheet.dimensions();
    {
        let mut data = Vec::new();
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&max_row.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&max_col.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        record(out, DIMENSION, &data);
    }
    // COLINFO — Java HSSF ColumnInfoRecord / sheet.setColumnWidth
    for (&col, &width) in &sheet.column_widths {
        record(out, COLINFO, &pack_colinfo(col, col, width, XF_GENERAL));
    }
    // ROW — Java HSSF RowRecord / setHeightInPoints
    let last_col_exclusive = u8::try_from(max_col.min(256)).unwrap_or(0);
    for (&row, &height) in &sheet.row_heights {
        record(out, ROW, &pack_row(row, 0, last_col_exclusive, height));
    }
    for (&(row, col), cell) in &sheet.cells {
        write_cell(out, row, col, cell, sst_index);
    }
    if !sheet.merges.is_empty() {
        let ranges: Vec<[u8; 8]> = sheet
            .merges
            .iter()
            .map(|m| {
                pack_merge_range(
                    m.first_row,
                    m.last_row,
                    u16::from(m.first_col),
                    u16::from(m.last_col),
                )
            })
            .collect();
        write_merge_cells(out, &ranges);
    }
    {
        let mut data = vec![0u8; 18];
        data[0] = 0xB6;
        data[1] = 0x06;
        record(out, WINDOW2, &data);
    }
    record(out, EOF, &[]);
}

fn write_cell(out: &mut Vec<u8>, row: u16, col: u8, cell: &Biff8Cell, sst_index: &HashMap<String, u32>) {
    match &cell.value {
        Biff8Value::Blank => write_blank(out, row, col, cell.xf),
        Biff8Value::Text(text) => {
            let idx = *sst_index.get(text).unwrap_or(&0);
            write_labelsst(out, row, col, cell.xf, idx);
        }
        Biff8Value::Number(n) => write_number(out, row, col, cell.xf, *n),
        Biff8Value::Bool(b) => write_boolerr(out, row, col, cell.xf, u8::from(*b), false),
    }
}

fn cell_header(data: &mut Vec<u8>, row: u16, col: u8, xf: u16) {
    data.extend_from_slice(&row.to_le_bytes());
    data.extend_from_slice(&u16::from(col).to_le_bytes());
    data.extend_from_slice(&xf.to_le_bytes());
}

fn write_blank(out: &mut Vec<u8>, row: u16, col: u8, xf: u16) {
    let mut data = Vec::new();
    cell_header(&mut data, row, col, xf);
    record(out, BLANK, &data);
}

fn write_number(out: &mut Vec<u8>, row: u16, col: u8, xf: u16, n: f64) {
    if let Some(rk) = encode_rk(n) {
        let mut data = Vec::new();
        cell_header(&mut data, row, col, xf);
        data.extend_from_slice(&rk.to_le_bytes());
        record(out, RK, &data);
    } else {
        let mut data = Vec::new();
        cell_header(&mut data, row, col, xf);
        data.extend_from_slice(&n.to_le_bytes());
        record(out, NUMBER, &data);
    }
}

fn write_labelsst(out: &mut Vec<u8>, row: u16, col: u8, xf: u16, sst: u32) {
    let mut data = Vec::new();
    cell_header(&mut data, row, col, xf);
    data.extend_from_slice(&sst.to_le_bytes());
    record(out, LABELSST, &data);
}

fn write_boolerr(out: &mut Vec<u8>, row: u16, col: u8, xf: u16, value: u8, is_error: bool) {
    let mut data = Vec::new();
    cell_header(&mut data, row, col, xf);
    data.push(value);
    data.push(u8::from(is_error));
    record(out, BOOLERR, &data);
}

