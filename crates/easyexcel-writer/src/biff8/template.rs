//! Minimal BIFF8 `.xls` template package (Java `withTemplate` / HSSF subset).
//!
//! # Approach
//!
//! Loads the OLE/CFB container, parses the `Workbook` stream into BIFF records,
//! and **preserves every untouched record byte-for-byte** (FONT / XF / SST /
//! MERGECELLS / existing cells). New values are inserted as inline `LABEL`
//! (0x0204) or `NUMBER` / `BOOLERR` / `BLANK` records immediately before the
//! target sheet's `EOF`, then `DIMENSION` and `BOUNDSHEET` stream offsets are
//! repaired. Other OLE streams (SummaryInformation, …) are kept by rewriting
//! only the `Workbook` / `Book` stream in place.
//!
//! # Java mapping
//!
//! | Java EasyExcel / POI | Rust |
//! |---|---|
//! | `EasyExcel.write(...).withTemplate(xls).sheet().doWrite(data)` | [`Biff8TemplatePackage`] + writer wiring |
//! | `HSSFWorkbook(templateStream)` | OLE open + Workbook record parse |
//! | `sheet.createRow(...).createCell(...).setCellValue(...)` | [`Biff8TemplatePackage::set_cell`] |
//! | POI keeps unedited records | unchanged BIFF records copied verbatim |
//!
//! # Still unsupported
//!
//! Placeholder `fill` (Java `ExcelWriter.fill` on POI `HSSFWorkbook`) remains
//! [`ExcelError::Unsupported`] at the template crate — list / `forceNewRow` /
//! horizontal fill need row insertion and SST mutation beyond this MVP.
//! Password-encrypted legacy workbooks are rejected.
//!
//! For `.xls` cell append (Java `withTemplate` + `doWrite`), use this package
//! via the writer facade instead of OOXML fill.

use std::io::{Cursor, Read, Write};
use std::path::Path;

use cfb::CompoundFile;
use easyexcel_core::{CellValue, ExcelError, Result};

use super::encode::{
    encode_rk, encode_unicode_string, BOOLERR, BOUNDSHEET, BOF, BLANK, DIMENSION, DT_WORKSHEET,
    EOF, LABEL, LABELSST, MAX_RECORD_DATA, NUMBER, RK, XF_GENERAL,
};
use super::{Biff8Cell, Biff8Value};

/// One framed BIFF record (`type` + payload).
#[derive(Debug, Clone)]
struct RawRecord {
    typ: u16,
    data: Vec<u8>,
}

/// Worksheet location inside the globals / sheet record list.
#[derive(Debug, Clone)]
struct SheetSpan {
    name: String,
    /// Index of the worksheet `BOF` record.
    bof_index: usize,
    /// Index of the worksheet `EOF` record (exclusive insert point is this index).
    eof_index: usize,
    /// Index of the `DIMENSION` record inside this sheet, when present.
    dimension_index: Option<usize>,
}

/// In-memory `.xls` template with record-preserving cell writes.
///
/// Corresponds to a loaded POI `HSSFWorkbook` used only for appending / overlay
/// cells while leaving the rest of the BIFF stream intact.
#[derive(Debug, Clone)]
pub struct Biff8TemplatePackage {
    /// Full OLE/CFB bytes (all streams); only `Workbook` is rewritten on save.
    ole_bytes: Vec<u8>,
    /// Workbook stream path (`Workbook` or `Book`).
    workbook_path: String,
    /// Parsed BIFF records from the Workbook stream.
    records: Vec<RawRecord>,
    /// Bound sheets in workbook order.
    sheets: Vec<SheetSpan>,
}

impl Biff8TemplatePackage {
    /// Loads an OLE `.xls` template from bytes.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Format`] when the bytes are not a readable BIFF8
    /// workbook, or [`ExcelError::Unsupported`] for empty / unusable templates.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if !bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
            return Err(ExcelError::Format(
                "xls template is not an OLE Compound File".to_owned(),
            ));
        }
        let (workbook_path, workbook) = read_workbook_stream(bytes)?;
        let records = split_records(&workbook)?;
        let sheets = discover_sheets(&records)?;
        if sheets.is_empty() {
            return Err(ExcelError::Format(
                "xls template Workbook contains no worksheets".to_owned(),
            ));
        }
        Ok(Self {
            ole_bytes: bytes.to_vec(),
            workbook_path,
            records,
            sheets,
        })
    }

    /// Loads an OLE `.xls` template from a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns I/O or format errors from [`Self::from_bytes`].
    pub fn from_path(path: &Path) -> Result<Self> {
        let bytes = std::fs::read(path).map_err(ExcelError::from)?;
        Self::from_bytes(&bytes)
    }

    /// Returns worksheet names in BoundSheet order.
    #[must_use]
    pub fn sheet_names(&self) -> Vec<String> {
        self.sheets.iter().map(|sheet| sheet.name.clone()).collect()
    }

    /// Returns the next zero-based append row for a sheet (Java `lastRowNum + 1`).
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::SheetNotFound`] when the sheet is absent.
    pub fn next_row_for_sheet(&self, sheet_name: &str) -> Result<u32> {
        let sheet = self.sheet(sheet_name)?;
        Ok(sheet_max_row(&self.records, sheet).map_or(0, |row| u32::from(row).saturating_add(1)))
    }

    /// Writes a cell value at `(row, col)`, replacing any existing cell record.
    ///
    /// Existing XF indexes are reused when overwriting a cell; new cells use
    /// [`XF_GENERAL`]. Unrelated records are left untouched.
    ///
    /// # Errors
    ///
    /// Returns format errors for out-of-range coordinates or unsupported values.
    pub fn set_cell(&mut self, sheet_name: &str, row: u32, col: usize, cell: Biff8Cell) -> Result<()> {
        let row = u16::try_from(row).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 65536 rows".to_owned())
        })?;
        let col = u8::try_from(col).map_err(|_| {
            ExcelError::Format("BIFF8 supports at most 256 columns".to_owned())
        })?;
        let sheet_index = self.sheet_index(sheet_name)?;
        let sheet = self.sheets[sheet_index].clone();
        let existing = find_cell_record(&self.records, &sheet, row, col);
        let xf = if let Some(index) = existing {
            // Preserve the template cell's XF (styles) when overwriting a value.
            if self.records[index].data.len() >= 6 {
                u16::from_le_bytes([
                    self.records[index].data[4],
                    self.records[index].data[5],
                ])
            } else {
                cell.xf
            }
        } else {
            cell.xf
        };
        let payload = encode_cell_record(row, col, xf, &cell.value)?;
        if let Some(index) = existing {
            self.records[index] = payload;
        } else {
            let insert_at = self.sheets[sheet_index].eof_index;
            self.records.insert(insert_at, payload);
            self.adjust_indices_after_insert(sheet_index, insert_at);
        }
        self.refresh_dimension(sheet_index)?;
        Ok(())
    }

    /// Convenience: writes a [`CellValue`] using general XF for new cells.
    ///
    /// # Errors
    ///
    /// Returns conversion or format errors from [`Self::set_cell`].
    pub fn set_cell_value(
        &mut self,
        sheet_name: &str,
        row: u32,
        col: usize,
        value: &CellValue,
    ) -> Result<()> {
        let cell = cell_value_to_template_cell(value)?;
        self.set_cell(sheet_name, row, col, cell)
    }

    /// Appends sparse rows starting at the sheet's next free row.
    ///
    /// # Errors
    ///
    /// Returns sheet / coordinate / value errors.
    pub fn append_rows(
        &mut self,
        sheet_name: &str,
        rows: &[Vec<(usize, CellValue)>],
    ) -> Result<u32> {
        let mut next = self.next_row_for_sheet(sheet_name)?;
        for row_values in rows {
            for (col, value) in row_values {
                self.set_cell_value(sheet_name, next, *col, value)?;
            }
            next = next.saturating_add(1);
        }
        Ok(next)
    }

    /// Serializes the package back to OLE/CFB bytes.
    ///
    /// # Errors
    ///
    /// Returns format or I/O errors when the Workbook stream cannot be rewritten.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let workbook = assemble_workbook(&self.records)?;
        rewrite_workbook_stream(&self.ole_bytes, &self.workbook_path, &workbook)
    }

    /// Returns all cell placeholders (`{key}` patterns) found in
    /// LABEL/LABELSST records across the workbook.
    ///
    /// Each entry is `(sheet_name, row, col, placeholder_text)`.
    /// The placeholder_text is the raw BIFF8 string content including
    /// the `{` and `}` delimiters.
    #[must_use]
    pub fn scan_placeholders(&self) -> Vec<(String, u16, u8, String)> {
        let mut placeholders = Vec::new();
        for sheet in &self.sheets {
            for (idx, record) in self.records.iter().enumerate() {
                if idx < sheet.bof_index || idx >= sheet.eof_index {
                    continue;
                }
        let (row, col, text) = match record.typ {
            LABEL => decode_label_payload(&record.data),
            LABELSST => decode_labelsst_payload(&record.data),
            _ => continue,
        };
        if let Some(ref text) = text {
            if text.contains('{') && text.contains('}') {
                placeholders.push((sheet.name.clone(), row, col, text.clone()));
            }
        }
            }
        }
        placeholders
    }

    /// Replaces a cell value at `(row, col)` on the given sheet with
    /// a new BIFF8 LABEL record containing the replacement text.
    ///
    /// # Errors
    ///
    /// Returns format errors for out-of-range coordinates.
    pub fn replace_label(
        &mut self,
        sheet_name: &str,
        row: u16,
        col: u8,
        replacement: &str,
    ) -> Result<()> {
        let sheet_index = self.sheet_index(sheet_name)?;
        let sheet = &self.sheets[sheet_index];
    let existing = find_cell_record(&self.records, sheet, row, col);
    let xf = if let Some(index) = existing {
        if self.records[index].data.len() >= 6 {
            u16::from_le_bytes([self.records[index].data[4], self.records[index].data[5]])
        } else {
            XF_GENERAL
        }
    } else {
        XF_GENERAL
    };
        let cell = Biff8Cell {
            value: Biff8Value::Text(replacement.to_owned()),
            xf,
        };
        let payload = encode_cell_record(row, col, xf, &cell.value)?;
        if let Some(index) = existing {
            self.records[index] = payload;
        } else {
            let insert_at = self.sheets[sheet_index].eof_index;
            self.records.insert(insert_at, payload);
            self.adjust_indices_after_insert(sheet_index, insert_at);
        }
        self.refresh_dimension(sheet_index)?;
        Ok(())
    }

    /// Writes the package to a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns I/O or format errors.
    pub fn save_to_path(&self, path: &Path) -> Result<()> {
        let bytes = self.to_bytes()?;
        if let Some(parent) = path.parent()
            && !parent.as_os_str().is_empty()
        {
            std::fs::create_dir_all(parent)?;
        }
        std::fs::write(path, bytes).map_err(ExcelError::from)
    }

    /// Writes the package to an arbitrary writer.
    ///
    /// # Errors
    ///
    /// Returns I/O or format errors.
    pub fn save_to_writer(&self, output: &mut dyn Write) -> Result<()> {
        let bytes = self.to_bytes()?;
        output.write_all(&bytes)?;
        output.flush()?;
        Ok(())
    }

    fn sheet(&self, name: &str) -> Result<&SheetSpan> {
        self.sheets
            .iter()
            .find(|sheet| sheet.name == name)
            .ok_or_else(|| ExcelError::SheetNotFound(name.to_owned()))
    }

    fn sheet_index(&self, name: &str) -> Result<usize> {
        self.sheets
            .iter()
            .position(|sheet| sheet.name == name)
            .ok_or_else(|| ExcelError::SheetNotFound(name.to_owned()))
    }

    /// After inserting a record at `insert_at`, shift later sheet indices.
    fn adjust_indices_after_insert(&mut self, sheet_index: usize, insert_at: usize) {
        for (index, sheet) in self.sheets.iter_mut().enumerate() {
            if sheet.bof_index >= insert_at {
                sheet.bof_index += 1;
            }
            if sheet.eof_index >= insert_at {
                sheet.eof_index += 1;
            }
            if let Some(dim) = sheet.dimension_index.as_mut()
                && *dim >= insert_at
            {
                *dim += 1;
            }
            if index == sheet_index {
                // Insert is always before EOF of this sheet.
                debug_assert!(sheet.eof_index > insert_at || sheet.eof_index == insert_at + 1);
            }
        }
    }

    fn refresh_dimension(&mut self, sheet_index: usize) -> Result<()> {
        let sheet = self.sheets[sheet_index].clone();
        let (max_row, max_col) = sheet_dimensions(&self.records, &sheet);
        let mut data = Vec::with_capacity(14);
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&u32::from(max_row).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&u16::from(max_col).to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        if let Some(dim_index) = sheet.dimension_index {
            self.records[dim_index] = RawRecord {
                typ: DIMENSION,
                data,
            };
        } else {
            let insert_at = sheet.bof_index + 1;
            self.records.insert(
                insert_at,
                RawRecord {
                    typ: DIMENSION,
                    data,
                },
            );
            self.sheets[sheet_index].dimension_index = Some(insert_at);
            self.adjust_indices_after_insert(sheet_index, insert_at);
        }
        Ok(())
    }
}

/// Converts a high-level cell value into a BIFF template cell.
fn cell_value_to_template_cell(value: &CellValue) -> Result<Biff8Cell> {
    let mapped = match value {
        CellValue::Empty => Biff8Value::Blank,
        CellValue::String(text)
        | CellValue::Error(text)
        | CellValue::Hyperlink { text, .. }
        | CellValue::Formula(text) => Biff8Value::Text(text.clone()),
        CellValue::RichText(rich) => Biff8Value::Text(rich.text_string().to_owned()),
        CellValue::Bool(flag) => Biff8Value::Bool(*flag),
        CellValue::Int(number) => Biff8Value::Number(*number as f64),
        CellValue::Float(number) => Biff8Value::Number(*number),
        CellValue::Decimal(number) => Biff8Value::Number(number.to_string().parse().unwrap_or(0.0)),
        CellValue::Date(date) => {
            return Ok(Biff8Cell::date_serial(super::date_to_excel_serial(*date)));
        }
        CellValue::DateTime(datetime) => {
            return Ok(Biff8Cell::datetime_serial(super::datetime_to_excel_serial(
                *datetime,
            )));
        }
        CellValue::Comment { value, .. } | CellValue::Images { value, .. } => {
            return cell_value_to_template_cell(value);
        }
        CellValue::Image(_) => {
            return Err(ExcelError::Unsupported(
                "legacy XLS writing does not support images".to_owned(),
            ));
        }
    };
    Ok(Biff8Cell::general(mapped))
}

fn encode_cell_record(row: u16, col: u8, xf: u16, value: &Biff8Value) -> Result<RawRecord> {
    let mut data = Vec::new();
    data.extend_from_slice(&row.to_le_bytes());
    data.extend_from_slice(&u16::from(col).to_le_bytes());
    data.extend_from_slice(&xf.to_le_bytes());
    match value {
        Biff8Value::Blank => Ok(RawRecord {
            typ: BLANK,
            data,
        }),
        Biff8Value::Bool(flag) => {
            data.push(u8::from(*flag));
            data.push(0);
            Ok(RawRecord {
                typ: BOOLERR,
                data,
            })
        }
        Biff8Value::Number(number) => {
            if let Some(rk) = encode_rk(*number) {
                data.extend_from_slice(&rk.to_le_bytes());
                Ok(RawRecord { typ: RK, data })
            } else {
                data.extend_from_slice(&number.to_le_bytes());
                Ok(RawRecord {
                    typ: NUMBER,
                    data,
                })
            }
        }
        Biff8Value::Text(text) => {
            // Inline LABEL avoids mutating the template SST (preserves indices).
            let encoded = encode_unicode_string(text);
            if data.len() + encoded.len() > MAX_RECORD_DATA {
                return Err(ExcelError::Format(
                    "xls template LABEL cell exceeds BIFF record size".to_owned(),
                ));
            }
            data.extend_from_slice(&encoded);
            Ok(RawRecord {
                typ: LABEL,
                data,
            })
        }
    }
}

fn read_workbook_stream(bytes: &[u8]) -> Result<(String, Vec<u8>)> {
    let cursor = Cursor::new(bytes.to_vec());
    let mut cf = CompoundFile::open(cursor).map_err(|error| {
        ExcelError::Format(format!("cannot open xls OLE container: {error}"))
    })?;
    for path in ["/Workbook", "/Book", "Workbook", "Book"] {
        if cf.is_stream(path) {
            let mut stream = cf.open_stream(path).map_err(|error| {
                ExcelError::Format(format!("cannot open {path} stream: {error}"))
            })?;
            let mut workbook = Vec::new();
            stream.read_to_end(&mut workbook)?;
            let normalized = if path.ends_with("Book") && !path.ends_with("Workbook") {
                "Book"
            } else {
                "Workbook"
            };
            return Ok((normalized.to_owned(), workbook));
        }
    }
    Err(ExcelError::Format(
        "xls template missing Workbook/Book stream".to_owned(),
    ))
}

fn rewrite_workbook_stream(
    ole_bytes: &[u8],
    workbook_path: &str,
    workbook: &[u8],
) -> Result<Vec<u8>> {
    let mut cursor = Cursor::new(ole_bytes.to_vec());
    {
        let mut cf = CompoundFile::open(&mut cursor).map_err(|error| {
            ExcelError::Format(format!("cannot reopen xls OLE container: {error}"))
        })?;
        {
            let mut stream = cf.open_stream(workbook_path).map_err(|error| {
                ExcelError::Format(format!("cannot rewrite {workbook_path}: {error}"))
            })?;
            stream.set_len(0).map_err(|error| {
                ExcelError::Format(format!("cannot truncate {workbook_path}: {error}"))
            })?;
            stream.write_all(workbook)?;
            stream.flush()?;
        }
        cf.flush()
            .map_err(|error| ExcelError::Format(format!("cannot flush OLE container: {error}")))?;
    }
    Ok(cursor.into_inner())
}

fn split_records(workbook: &[u8]) -> Result<Vec<RawRecord>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= workbook.len() {
        let typ = u16::from_le_bytes([workbook[offset], workbook[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([
            workbook[offset + 2],
            workbook[offset + 3],
        ]));
        offset += 4;
        if offset + length > workbook.len() {
            return Err(ExcelError::Format(format!(
                "truncated BIFF record type=0x{typ:04X} len={length}"
            )));
        }
        records.push(RawRecord {
            typ,
            data: workbook[offset..offset + length].to_vec(),
        });
        offset += length;
    }
    if records.is_empty() {
        return Err(ExcelError::Format(
            "xls template Workbook stream has no BIFF records".to_owned(),
        ));
    }
    Ok(records)
}

fn discover_sheets(records: &[RawRecord]) -> Result<Vec<SheetSpan>> {
    let mut names = Vec::new();
    for record in records {
        if record.typ == BOUNDSHEET {
            names.push(decode_boundsheet_name(&record.data)?);
        }
    }
    let mut sheets = Vec::new();
    let mut name_iter = names.into_iter();
    let mut index = 0usize;
    while index < records.len() {
        let record = &records[index];
        if record.typ == BOF && is_worksheet_bof(&record.data) {
            let name = name_iter.next().unwrap_or_else(|| format!("Sheet{}", sheets.len() + 1));
            let bof_index = index;
            let mut dimension_index = None;
            let mut eof_index = None;
            index += 1;
            while index < records.len() {
                match records[index].typ {
                    DIMENSION if dimension_index.is_none() => dimension_index = Some(index),
                    EOF => {
                        eof_index = Some(index);
                        break;
                    }
                    BOF => {
                        return Err(ExcelError::Format(
                            "xls template has nested worksheet BOF without EOF".to_owned(),
                        ));
                    }
                    _ => {}
                }
                index += 1;
            }
            let eof_index = eof_index.ok_or_else(|| {
                ExcelError::Format(format!("xls template sheet `{name}` is missing EOF"))
            })?;
            sheets.push(SheetSpan {
                name,
                bof_index,
                eof_index,
                dimension_index,
            });
        }
        index += 1;
    }
    Ok(sheets)
}

fn is_worksheet_bof(data: &[u8]) -> bool {
    data.len() >= 4 && u16::from_le_bytes([data[2], data[3]]) == DT_WORKSHEET
}

fn decode_boundsheet_name(data: &[u8]) -> Result<String> {
    // lbPlyPos(4) + hsState(1) + dt(1) + short XLUnicodeString
    if data.len() < 8 {
        return Err(ExcelError::Format(
            "BOUNDSHEET record is too short".to_owned(),
        ));
    }
    let cch = usize::from(data[6]);
    let compressed = data[7] & 0x01 == 0;
    let raw = &data[8..];
    if compressed {
        let take = cch.min(raw.len());
        Ok(raw[..take].iter().map(|&byte| char::from(byte)).collect())
    } else {
        let take = cch.saturating_mul(2).min(raw.len());
        let units: Vec<u16> = raw[..take]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        Ok(String::from_utf16_lossy(&units))
    }
    }

/// Decodes a BIFF8 LABEL record payload, returning `(row, col, text)`.
fn decode_label_payload(data: &[u8]) -> (u16, u8, Option<String>) {
    if data.len() < 8 {
        return (0, 0, None);
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    // Bytes 4-5 are XF index; bytes 6-7 are string length
    let byte_len = u16::from_le_bytes([data[6], data[7]]) as usize;
    let string_data = &data[8..];
    let text = if string_data.len() >= byte_len + 1 {
        // BIFF8 short string: leading byte is length, followed by bytes
        let len = string_data[0] as usize;
        if len <= string_data.len().saturating_sub(1)
            && len <= byte_len
        {
            let raw = &string_data[1..=len.min(string_data.len().saturating_sub(1))];
            String::from_utf8_lossy(raw).into_owned()
        } else {
            String::new()
        }
    } else {
        String::new()
    };
    (row, col as u8, if text.is_empty() { None } else { Some(text) })
}

/// Decodes a BIFF8 LABELSST record payload, returning `(row, col, text)`.
/// LABELSST references the Shared String Table — since we don't have
/// the SST available here, we return None for the text and let the
/// caller handle SST lookups separately.
fn decode_labelsst_payload(data: &[u8]) -> (u16, u8, Option<String>) {
    if data.len() < 8 {
        return (0, 0, None);
    }
    let row = u16::from_le_bytes([data[0], data[1]]);
    let col = u16::from_le_bytes([data[2], data[3]]);
    // Bytes 4-5: XF, bytes 6-9: SST index (u32)
    if data.len() >= 10 {
        let _sst_index = u32::from_le_bytes([data[6], data[7], data[8], data[9]]);
        // SST-based records can't be decoded without the shared string table;
        // caller should use LABEL records for placeholder detection.
        (row, col as u8, None)
    } else {
        (row, col as u8, None)
    }
}

fn sheet_max_row(records: &[RawRecord], sheet: &SheetSpan) -> Option<u16> {
    let mut maximum = None;
    for record in &records[sheet.bof_index..=sheet.eof_index] {
        if let Some((row, _)) = cell_coords(record) {
            maximum = Some(maximum.map_or(row, |current: u16| current.max(row)));
        }
    }
    maximum
}

fn sheet_dimensions(records: &[RawRecord], sheet: &SheetSpan) -> (u16, u8) {
    let mut max_row = 0u16;
    let mut max_col = 0u8;
    for record in &records[sheet.bof_index..=sheet.eof_index] {
        if let Some((row, col)) = cell_coords(record) {
            max_row = max_row.max(row.saturating_add(1));
            max_col = max_col.max(col.saturating_add(1));
        }
    }
    (max_row, max_col)
}

fn cell_coords(record: &RawRecord) -> Option<(u16, u8)> {
    match record.typ {
        LABEL | LABELSST | NUMBER | RK | BOOLERR | BLANK => {
            if record.data.len() < 4 {
                return None;
            }
            let row = u16::from_le_bytes([record.data[0], record.data[1]]);
            let col = u16::from_le_bytes([record.data[2], record.data[3]]);
            let col = u8::try_from(col).ok()?;
            Some((row, col))
        }
        _ => None,
    }
}

fn find_cell_record(
    records: &[RawRecord],
    sheet: &SheetSpan,
    row: u16,
    col: u8,
) -> Option<usize> {
    for index in sheet.bof_index..=sheet.eof_index {
        if cell_coords(&records[index]) == Some((row, col)) {
            return Some(index);
        }
    }
    None
}

fn assemble_workbook(records: &[RawRecord]) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    let mut boundsheet_patches = Vec::new();
    let mut sheet_offsets = Vec::new();
    for record in records {
        if record.typ == BOUNDSHEET {
            // Patch site: absolute offset of lbPlyPos inside the assembled stream.
            boundsheet_patches.push(out.len() + 4);
        }
        if record.typ == BOF && is_worksheet_bof(&record.data) {
            sheet_offsets.push(out.len() as u32);
        }
        write_raw_record(&mut out, record)?;
    }
    if boundsheet_patches.len() != sheet_offsets.len() {
        return Err(ExcelError::Format(format!(
            "BOUNDSHEET count ({}) does not match worksheet BOF count ({})",
            boundsheet_patches.len(),
            sheet_offsets.len()
        )));
    }
    for (patch_at, offset) in boundsheet_patches.into_iter().zip(sheet_offsets) {
        out[patch_at..patch_at + 4].copy_from_slice(&offset.to_le_bytes());
    }
    Ok(out)
}

fn write_raw_record(out: &mut Vec<u8>, record: &RawRecord) -> Result<()> {
    if record.data.len() > MAX_RECORD_DATA {
        return Err(ExcelError::Format(format!(
            "BIFF record 0x{:04X} payload exceeds {MAX_RECORD_DATA} bytes",
            record.typ
        )));
    }
    out.extend_from_slice(&record.typ.to_le_bytes());
    out.extend_from_slice(&(record.data.len() as u16).to_le_bytes());
    out.extend_from_slice(&record.data);
    Ok(())
}

/// Returns whether `bytes` look like an OLE `.xls` compound document.
#[must_use]
pub fn looks_like_xls(bytes: &[u8]) -> bool {
    bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0])
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::biff8::{Biff8Book, Biff8Cell, Biff8Value};

    #[test]
    fn loads_self_written_template_writes_cell_and_preserves_other_cells() {
        let mut book = Biff8Book::default();
        {
            let sheet = book.sheet_mut("Sheet1");
            sheet
                .set(0, 0, Biff8Cell::general(Biff8Value::Text("keep-a".into())))
                .unwrap();
            sheet
                .set(0, 1, Biff8Cell::general(Biff8Value::Text("keep-b".into())))
                .unwrap();
            sheet
                .set(1, 0, Biff8Cell::general(Biff8Value::Number(11.0)))
                .unwrap();
        }
        let template_bytes = book.to_cfb_bytes().unwrap();

        let mut package = Biff8TemplatePackage::from_bytes(&template_bytes).unwrap();
        assert_eq!(package.sheet_names(), vec!["Sheet1".to_owned()]);
        assert_eq!(package.next_row_for_sheet("Sheet1").unwrap(), 2);

        // Overwrite one cell and append another — other cells must survive.
        package
            .set_cell(
                "Sheet1",
                1,
                0,
                Biff8Cell::general(Biff8Value::Number(99.0)),
            )
            .unwrap();
        package
            .set_cell(
                "Sheet1",
                2,
                0,
                Biff8Cell::general(Biff8Value::Text("appended".into())),
            )
            .unwrap();

        let out = package.to_bytes().unwrap();
        // Unmodified SST / early records: original Workbook should still contain
        // the shared strings for keep-a / keep-b (LABELSST path from Biff8Book).
        let (_, original_wb) = read_workbook_stream(&template_bytes).unwrap();
        let (_, updated_wb) = read_workbook_stream(&out).unwrap();
        assert!(
            updated_wb.len() >= original_wb.len(),
            "appending cells should not shrink the Workbook stream"
        );

        // Round-trip via calamine to assert values.
        use calamine::{Data, DataType, Reader, Xls, open_workbook_from_rs};
        let mut xls: Xls<_> = open_workbook_from_rs(Cursor::new(out)).unwrap();
        let range = xls.worksheet_range("Sheet1").unwrap();
        assert_eq!(
            range.get((0, 0)).and_then(DataType::as_string).as_deref(),
            Some("keep-a")
        );
        assert_eq!(
            range.get((0, 1)).and_then(DataType::as_string).as_deref(),
            Some("keep-b")
        );
        // NUMBER/RK may surface as Int or Float depending on calamine decoding.
        let cell = range.get((1, 0));
        assert!(
            matches!(cell, Some(Data::Float(v)) if *v == 99.0)
                || matches!(cell, Some(Data::Int(99))),
            "expected numeric 99, got {cell:?}"
        );
        assert_eq!(
            range.get((2, 0)).and_then(DataType::as_string).as_deref(),
            Some("appended")
        );
    }

    #[test]
    fn preserves_unmodified_record_bytes_when_appending() {
        let mut book = Biff8Book::default();
        book.sheet_mut("Sheet1")
            .set(0, 0, Biff8Cell::general(Biff8Value::Text("header".into())))
            .unwrap();
        let template_bytes = book.to_cfb_bytes().unwrap();
        let mut package = Biff8TemplatePackage::from_bytes(&template_bytes).unwrap();
        let before = package.records.clone();
        package
            .append_rows(
                "Sheet1",
                &[vec![(0usize, CellValue::String("row1".into()))]],
            )
            .unwrap();

        // Every record that existed before must still appear with identical type+payload
        // except DIMENSION (updated) and the newly inserted LABEL before EOF.
        let mut matched = 0usize;
        for original in &before {
            if original.typ == DIMENSION {
                continue;
            }
            assert!(
                package
                    .records
                    .iter()
                    .any(|record| record.typ == original.typ && record.data == original.data),
                "missing preserved record typ=0x{:04X}",
                original.typ
            );
            matched += 1;
        }
        assert!(matched > 0);
        assert!(package.records.iter().any(|record| record.typ == LABEL));
    }
}
