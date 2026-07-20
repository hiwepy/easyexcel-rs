//! Template-backed workbook seeding for Java `withTemplate` + `doWrite`.
//!
//! **Default path:** ZIP/OOXML preserve ([`TemplatePackage`]). Clone the template
//! package, keep `xl/styles.xml` and worksheet `mergeCells` intact, append typed
//! rows into `sheetData`, and when a requested sheet is missing create a new
//! worksheet part without rewriting existing sheets.
//!
//! **Legacy path:** calamine → `rust_xlsxwriter` value replay. Styles, merges,
//! images, comments, drawings, and column widths are **not** preserved. Used
//! only when callers explicitly set
//! [`crate::WriteOptions::use_legacy_template_seed`].

use std::fmt::Write as _;
use std::io::{Cursor, Read, Seek, Write};
use std::path::Path;

use calamine::{Data, DataType, Reader, Xlsx, open_workbook_from_rs};
use easyexcel_core::{CellValue, ExcelError, Result};
use rust_xlsxwriter::{Format, Workbook, Worksheet};
use zip::CompressionMethod;
use zip::read::ZipArchive;
use zip::write::{SimpleFileOptions, ZipWriter};

use crate::format_error;

/// One worksheet loaded from a template package (value snapshot for legacy seed).
#[derive(Debug, Clone)]
pub(crate) struct TemplateSheetData {
    /// Worksheet name from the template workbook.
    pub name: String,
    /// Non-empty cells as `(row, column, value)` with zero-based coordinates.
    pub cells: Vec<(u32, u16, Data)>,
    /// Next zero-based row index available for append (Java `getNewRowIndexAndStartDoWrite`).
    pub next_row: u32,
}

/// One ZIP entry retained from a template XLSX package.
#[derive(Debug, Clone)]
pub(crate) struct TemplateZipEntry {
    /// Entry path inside the OOXML package.
    pub name: String,
    /// Whether this entry is a directory marker.
    pub is_dir: bool,
    /// Compression method copied from the template.
    pub compression: CompressionMethod,
    /// Optional UNIX mode bits from the template.
    pub unix_mode: Option<u32>,
    /// Raw entry bytes (empty for directories).
    pub bytes: Vec<u8>,
}

/// In-memory XLSX template package used by the ZIP preserve write path.
#[derive(Debug, Clone)]
pub(crate) struct TemplatePackage {
    entries: Vec<TemplateZipEntry>,
}

impl TemplatePackage {
    /// Loads an XLSX template package from bytes.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Format`] when the bytes are not a readable ZIP/OOXML package.
    pub(crate) fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let mut archive = ZipArchive::new(Cursor::new(bytes.to_vec())).map_err(format_error)?;
        let mut entries = Vec::with_capacity(archive.len());
        for index in 0..archive.len() {
            let mut entry = archive.by_index(index).map_err(format_error)?;
            let mut bytes = Vec::new();
            if !entry.is_dir() {
                entry.read_to_end(&mut bytes)?;
            }
            entries.push(TemplateZipEntry {
                name: entry.name().to_owned(),
                is_dir: entry.is_dir(),
                compression: entry.compression(),
                unix_mode: entry.unix_mode(),
                bytes,
            });
        }
        Ok(Self { entries })
    }

    /// Returns worksheet names in workbook order.
    ///
    /// # Errors
    ///
    /// Returns a format error when workbook metadata cannot be parsed.
    pub(crate) fn sheet_names(&self) -> Result<Vec<String>> {
        Ok(self.workbook_sheets()?.into_iter().map(|(name, _)| name).collect())
    }

    /// Returns the next zero-based append row for a worksheet name.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::SheetNotFound`] when the sheet is absent.
    pub(crate) fn next_row_for_sheet(&self, sheet_name: &str) -> Result<u32> {
        let path = self.worksheet_path_by_name(sheet_name)?;
        let xml = self.entry_xml(&path)?;
        let max = worksheet_max_row(&xml);
        // Java WriteSheetHolder.TEMPLATE_EMPTY: lastRowNum + 1 when any row exists.
        if max == 0 && !xml.contains("<row") {
            Ok(0)
        } else {
            Ok(u32::try_from(max.saturating_add(1)).unwrap_or(u32::MAX))
        }
    }

    /// Resolves the worksheet part path for a sheet name.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::SheetNotFound`] when the sheet is absent.
    pub(crate) fn worksheet_path_by_name(&self, sheet_name: &str) -> Result<String> {
        let sheets = self.workbook_sheets()?;
        let selected = sheets
            .iter()
            .find(|(name, _)| name == sheet_name)
            .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_owned()))?;
        self.worksheet_part_for_relationship(&selected.1, &selected.0)
    }

    /// Resolves the worksheet part path for a zero-based sheet index.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::SheetNotFound`] when the index is out of range.
    pub(crate) fn worksheet_path_by_index(&self, index: usize) -> Result<(String, String)> {
        let sheets = self.workbook_sheets()?;
        let selected = sheets.get(index).ok_or_else(|| {
            ExcelError::SheetNotFound(format!("sheet index {index}"))
        })?;
        let path = self.worksheet_part_for_relationship(&selected.1, &selected.0)?;
        Ok((selected.0.clone(), path))
    }

    /// Ensures a worksheet exists; creates an empty one when the name is new.
    ///
    /// Existing worksheets, `xl/styles.xml`, and their `mergeCells` are left
    /// untouched. Mirrors Java creating a sheet that is absent from the template.
    ///
    /// # Errors
    ///
    /// Returns a format error when workbook / relationship metadata cannot be updated.
    pub(crate) fn ensure_sheet(&mut self, sheet_name: &str) -> Result<()> {
        if self
            .sheet_names()?
            .iter()
            .any(|name| name == sheet_name)
        {
            return Ok(());
        }
        self.create_sheet(sheet_name)
    }

    /// Creates an empty worksheet part and registers it in the package.
    ///
    /// Existing worksheets, `xl/styles.xml`, and their `mergeCells` stay
    /// untouched. The new sheet inherits `sheetFormatPr` / `cols` from the first
    /// template sheet when present (workbook styles remain shared).
    ///
    /// # Errors
    ///
    /// Returns a format error when workbook / relationship metadata cannot be updated.
    pub(crate) fn create_sheet(&mut self, sheet_name: &str) -> Result<()> {
        let sheet_part = next_worksheet_part_name(&self.entries);
        let workbook_index = self
            .entries
            .iter()
            .position(|entry| entry.name.eq_ignore_ascii_case("xl/workbook.xml"))
            .ok_or_else(|| ExcelError::Format("template missing xl/workbook.xml".to_owned()))?;
        let rels_index = self
            .entries
            .iter()
            .position(|entry| {
                entry
                    .name
                    .eq_ignore_ascii_case("xl/_rels/workbook.xml.rels")
            })
            .ok_or_else(|| {
                ExcelError::Format("template missing xl/_rels/workbook.xml.rels".to_owned())
            })?;
        let content_types_index = self
            .entries
            .iter()
            .position(|entry| entry.name.eq_ignore_ascii_case("[Content_Types].xml"))
            .ok_or_else(|| ExcelError::Format("template missing [Content_Types].xml".to_owned()))?;

        let workbook_xml = String::from_utf8(self.entries[workbook_index].bytes.clone())
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let rels_xml = String::from_utf8(self.entries[rels_index].bytes.clone())
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let content_types_xml = String::from_utf8(self.entries[content_types_index].bytes.clone())
            .map_err(|error| ExcelError::Format(error.to_string()))?;

        let relationship_id = next_relationship_id(&rels_xml);
        let sheet_id = next_sheet_id(&workbook_xml);
        let escaped_name = escape_xml(sheet_name);
        let sheet_tag = format!(
            "<sheet name=\"{escaped_name}\" sheetId=\"{sheet_id}\" r:id=\"{relationship_id}\"/>"
        );
        let updated_workbook = insert_before_close_tag(&workbook_xml, "</sheets>", &sheet_tag)?;
        let relationship_tag = format!(
            "<Relationship Id=\"{relationship_id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/{}\"/>",
            sheet_part
                .strip_prefix("xl/worksheets/")
                .unwrap_or(sheet_part.as_str())
        );
        let updated_rels =
            insert_before_close_tag(&rels_xml, "</Relationships>", &relationship_tag)?;
        let override_tag = format!(
            "<Override PartName=\"/{sheet_part}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        );
        let updated_types =
            insert_before_close_tag(&content_types_xml, "</Types>", &override_tag)?;

        self.entries[workbook_index].bytes = updated_workbook.into_bytes();
        self.entries[rels_index].bytes = updated_rels.into_bytes();
        self.entries[content_types_index].bytes = updated_types.into_bytes();
        // Inherit sheet-level format (row height / column widths) from the first
        // template sheet when present; never copy its cells, styles indexes, or merges.
        let worksheet_bytes = blank_worksheet_with_inherited_format(&self.entries);
        self.entries.push(TemplateZipEntry {
            name: sheet_part,
            is_dir: false,
            compression: CompressionMethod::Deflated,
            unix_mode: None,
            bytes: worksheet_bytes,
        });
        Ok(())
    }

    /// Appends typed rows into a worksheet's `sheetData` without rewriting styles/merges.
    ///
    /// New cells are written as inline values and do not receive template `s=` indexes.
    /// Existing `xl/styles.xml` and `mergeCells` parts are left untouched.
    ///
    /// # Errors
    ///
    /// Returns a format error when the worksheet XML cannot be updated.
    pub(crate) fn append_rows(
        &mut self,
        sheet_name: &str,
        rows: &[Vec<(usize, CellValue)>],
    ) -> Result<u32> {
        if rows.is_empty() {
            return self.next_row_for_sheet(sheet_name);
        }
        let path = self.worksheet_path_by_name(sheet_name)?;
        let entry = self
            .entries
            .iter_mut()
            .find(|entry| entry.name.eq_ignore_ascii_case(&path))
            .ok_or_else(|| ExcelError::Format(format!("template does not contain {path}")))?;
        let xml = String::from_utf8(std::mem::take(&mut entry.bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let (updated, next_row) = append_sparse_rows_to_xml(&xml, rows)?;
        entry.bytes = updated.into_bytes();
        Ok(next_row)
    }

    /// Serializes the package to owned XLSX bytes.
    ///
    /// # Errors
    ///
    /// Returns a format or I/O error when ZIP writing fails.
    pub(crate) fn to_bytes(&self) -> Result<Vec<u8>> {
        let cursor = Cursor::new(Vec::new());
        let finished = write_entries_to(Box::new(cursor), &self.entries)?;
        finished
            .into_inner()
            .map_err(|_| ExcelError::Format("ZIP output buffer type changed".to_owned()))
    }

    /// Writes the package to a filesystem path.
    ///
    /// # Errors
    ///
    /// Returns an I/O or format error.
    pub(crate) fn save_to_path(&self, path: &Path) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes).map_err(ExcelError::from)
    }

    /// Writes the package to an arbitrary writer.
    ///
    /// # Errors
    ///
    /// Returns an I/O or format error.
    pub(crate) fn save_to_writer(&self, output: &mut dyn Write) -> Result<()> {
        let bytes = self.to_bytes()?;
        output.write_all(&bytes)?;
        output.flush()?;
        Ok(())
    }

    fn workbook_sheets(&self) -> Result<Vec<(String, String)>> {
        let workbook = self
            .entries
            .iter()
            .find(|entry| entry.name.eq_ignore_ascii_case("xl/workbook.xml"))
            .ok_or_else(|| ExcelError::Format("template missing xl/workbook.xml".to_owned()))?;
        let xml = std::str::from_utf8(&workbook.bytes)
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        Ok(xml_elements(xml, "sheet")
            .filter_map(|element| {
                Some((
                    attribute_value(element, "name")?.to_owned(),
                    attribute_value(element, "r:id")?.to_owned(),
                ))
            })
            .collect())
    }

    fn worksheet_part_for_relationship(
        &self,
        relationship_id: &str,
        sheet_name: &str,
    ) -> Result<String> {
        let relationships = self
            .entries
            .iter()
            .find(|entry| {
                entry
                    .name
                    .eq_ignore_ascii_case("xl/_rels/workbook.xml.rels")
            })
            .ok_or_else(|| {
                ExcelError::Format("template missing xl/_rels/workbook.xml.rels".to_owned())
            })?;
        let xml = std::str::from_utf8(&relationships.bytes)
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let target = xml_elements(xml, "Relationship")
            .find(|element| attribute_value(element, "Id") == Some(relationship_id))
            .and_then(|element| attribute_value(element, "Target"))
            .ok_or_else(|| {
                ExcelError::Format(format!(
                    "workbook relationship {relationship_id} for sheet {sheet_name} is missing"
                ))
            })?;
        let normalized = normalize_workbook_target(target)?;
        self.entries
            .iter()
            .find(|entry| entry.name.eq_ignore_ascii_case(&normalized))
            .map(|entry| entry.name.clone())
            .ok_or_else(|| {
                ExcelError::Format(format!(
                    "worksheet part {normalized} for sheet {sheet_name} is missing"
                ))
            })
    }

    fn entry_xml(&self, path: &str) -> Result<String> {
        let entry = self
            .entries
            .iter()
            .find(|entry| entry.name.eq_ignore_ascii_case(path))
            .ok_or_else(|| ExcelError::Format(format!("template does not contain {path}")))?;
        String::from_utf8(entry.bytes.clone()).map_err(|error| ExcelError::Format(error.to_string()))
    }
}

/// Returns whether [`crate::WriteOptions`] carries a template source.
///
/// Corresponds to Java `WriteWorkbook.templateFile` / `templateInputStream`
/// being non-null.
#[must_use]
pub(crate) fn has_template(
    template_file: Option<&Path>,
    template_bytes: Option<&[u8]>,
) -> bool {
    template_file.is_some() || template_bytes.is_some()
}

/// Loads template bytes from a file path or an in-memory copy.
///
/// # Errors
///
/// Returns I/O errors when the template file cannot be read, or
/// [`ExcelError::Unsupported`] when no template source is configured.
pub(crate) fn load_template_bytes(
    template_file: Option<&Path>,
    template_bytes: Option<&[u8]>,
) -> Result<Vec<u8>> {
    if let Some(bytes) = template_bytes {
        return Ok(bytes.to_vec());
    }
    if let Some(path) = template_file {
        return std::fs::read(path).map_err(ExcelError::from);
    }
    Err(ExcelError::Unsupported(
        "with_template requires a template file or template bytes".to_owned(),
    ))
}

/// Rejects template types that Java also rejects for the XLSX ZIP path.
///
/// # Errors
///
/// - CSV templates → same as Java `ExcelGenerateException("csv cannot use template.")`
/// - XLS templates → rejected here for **XLSX output**; `.xls` output uses
///   [`crate::biff8::Biff8TemplatePackage`] instead (see writer `start` / `write_xls`).
pub(crate) fn validate_template_source(
    template_file: Option<&Path>,
    template_bytes: Option<&[u8]>,
) -> Result<()> {
    if let Some(path) = template_file {
        if is_csv_path(path) {
            return Err(ExcelError::Unsupported(
                "csv cannot use template.".to_owned(),
            ));
        }
        if is_xls_path(path) {
            return Err(ExcelError::Unsupported(
                "legacy XLS template cannot seed an XLSX workbook; write to a .xls path instead"
                    .to_owned(),
            ));
        }
    }
    if let Some(bytes) = template_bytes {
        if looks_like_csv(bytes) {
            return Err(ExcelError::Unsupported(
                "csv cannot use template.".to_owned(),
            ));
        }
        if looks_like_xls(bytes) {
            return Err(ExcelError::Unsupported(
                "legacy XLS template cannot seed an XLSX workbook; write to a .xls path instead"
                    .to_owned(),
            ));
        }
    }
    Ok(())
}

/// Parses an XLSX template into ordered sheet snapshots.
///
/// Used only by the explicit legacy value-replay path
/// ([`crate::WriteOptions::use_legacy_template_seed`]).
///
/// # Errors
///
/// Returns [`ExcelError::Format`] when the package is not a readable XLSX workbook.
pub(crate) fn load_template_sheets(bytes: &[u8]) -> Result<Vec<TemplateSheetData>> {
    let mut workbook: Xlsx<_> =
        open_workbook_from_rs(Cursor::new(bytes)).map_err(|error| {
            ExcelError::Format(format!("failed to open withTemplate workbook: {error}"))
        })?;
    let names = workbook.sheet_names().to_vec();
    if names.is_empty() {
        return Err(ExcelError::Format(
            "withTemplate workbook contains no worksheets".to_owned(),
        ));
    }
    let mut sheets = Vec::with_capacity(names.len());
    for name in names {
        let range = workbook.worksheet_range(&name).map_err(|error| {
            ExcelError::Format(format!(
                "failed to read withTemplate sheet `{name}`: {error}"
            ))
        })?;
        let mut cells = Vec::new();
        let mut last_row: Option<u32> = None;
        for (row, column, value) in range.used_cells() {
            if value.is_empty() {
                continue;
            }
            let row = u32::try_from(row).map_err(|_| {
                ExcelError::Format(format!(
                    "withTemplate sheet `{name}` row index {row} exceeds u32"
                ))
            })?;
            let column = u16::try_from(column).map_err(|_| {
                ExcelError::Format(format!(
                    "withTemplate sheet `{name}` column index {column} exceeds u16"
                ))
            })?;
            last_row = Some(last_row.map_or(row, |current| current.max(row)));
            cells.push((row, column, value.clone()));
        }
        // Java WriteSheetHolder.TEMPLATE_EMPTY: lastRowNum + 1 when any row exists.
        let next_row = last_row.map_or(0, |row| row.saturating_add(1));
        sheets.push(TemplateSheetData {
            name,
            cells,
            next_row,
        });
    }
    Ok(sheets)
}

/// Resolves the target sheet for Java `sheet()` / `sheet(no)` / `sheet(name)`.
///
/// Preference matches Java `WriteContextImpl.initSheet`:
/// 1. `sheet_index` when set
/// 2. otherwise match by `sheet_name`
/// 3. otherwise treat as a new sheet to create after template sheets
#[must_use]
pub(crate) fn resolve_template_target(
    sheets: &[TemplateSheetData],
    sheet_index: Option<usize>,
    sheet_name: &str,
) -> (usize, String, bool) {
    if let Some(index) = sheet_index {
        if let Some(sheet) = sheets.get(index) {
            return (index, sheet.name.clone(), false);
        }
        return (index, sheet_name.to_owned(), true);
    }
    if let Some((index, sheet)) = sheets
        .iter()
        .enumerate()
        .find(|(_, sheet)| sheet.name == sheet_name)
    {
        return (index, sheet.name.clone(), false);
    }
    (sheets.len(), sheet_name.to_owned(), true)
}

/// Resolves a template target against a ZIP package sheet list.
#[must_use]
pub(crate) fn resolve_package_target(
    sheet_names: &[String],
    sheet_index: Option<usize>,
    sheet_name: &str,
) -> (usize, String, bool) {
    if let Some(index) = sheet_index {
        if let Some(name) = sheet_names.get(index) {
            return (index, name.clone(), false);
        }
        return (index, sheet_name.to_owned(), true);
    }
    if let Some((index, name)) = sheet_names
        .iter()
        .enumerate()
        .find(|(_, name)| *name == sheet_name)
    {
        return (index, name.clone(), false);
    }
    (sheet_names.len(), sheet_name.to_owned(), true)
}

/// Writes loaded template sheets into a fresh `rust_xlsxwriter` workbook.
///
/// **Legacy only** ([`crate::WriteOptions::use_legacy_template_seed`]): values
/// only — styles/merges are not preserved. Prefer [`TemplatePackage`] by default.
///
/// # Errors
///
/// Returns worksheet naming or cell-write errors from `rust_xlsxwriter`.
pub(crate) fn seed_workbook_from_template(
    workbook: &mut Workbook,
    sheets: &[TemplateSheetData],
) -> Result<()> {
    for sheet in sheets {
        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet.name).map_err(format_error)?;
        for (row, column, value) in &sheet.cells {
            write_template_cell(worksheet, *row, *column, value)?;
        }
    }
    Ok(())
}

/// Writes a single calamine cell into a worksheet.
///
/// # Errors
///
/// Returns XLSX write errors from `rust_xlsxwriter`.
pub(crate) fn write_template_cell(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: &Data,
) -> Result<()> {
    match value {
        Data::Empty => Ok(()),
        Data::String(text) | Data::DateTimeIso(text) | Data::DurationIso(text) => worksheet
            .write_string(row, column, text)
            .map_err(format_error)
            .map(|_| ()),
        Data::Bool(flag) => worksheet
            .write_boolean(row, column, *flag)
            .map_err(format_error)
            .map(|_| ()),
        Data::Int(number) => {
            #[allow(clippy::cast_precision_loss)]
            let value = *number as f64;
            worksheet
                .write_number(row, column, value)
                .map_err(format_error)
                .map(|_| ())
        }
        Data::Float(number) => worksheet
            .write_number(row, column, *number)
            .map_err(format_error)
            .map(|_| ()),
        Data::DateTime(datetime) => {
            if let Some(chrono_value) = datetime.as_datetime() {
                let format = Format::new().set_num_format("yyyy-mm-dd hh:mm:ss");
                worksheet
                    .write_datetime_with_format(row, column, chrono_value, &format)
                    .map_err(format_error)
                    .map(|_| ())
            } else {
                worksheet
                    .write_number(row, column, datetime.as_f64())
                    .map_err(format_error)
                    .map(|_| ())
            }
        }
        Data::Error(error) => worksheet
            .write_string(row, column, error.to_string())
            .map_err(format_error)
            .map(|_| ()),
    }
}

fn append_sparse_rows_to_xml(
    xml: &str,
    rows: &[Vec<(usize, CellValue)>],
) -> Result<(String, u32)> {
    // Brand-new worksheets (and some Excel empties) use self-closing
    // `<sheetData/>`; expand so row append can splice before `</sheetData>`.
    let xml = expand_self_closing_sheet_data(xml)?;
    let Some(sheet_data_end) = xml.find("</sheetData>") else {
        return Err(ExcelError::Format(
            "worksheet does not contain sheetData".to_owned(),
        ));
    };
    let max_row = worksheet_max_row(&xml[..sheet_data_end]);
    let next_row = if max_row == 0 && !xml[..sheet_data_end].contains("<row") {
        1usize
    } else {
        max_row.saturating_add(1)
    };
    let mut appended = String::new();
    for (row_offset, values) in rows.iter().enumerate() {
        let row_index = next_row + row_offset;
        write!(appended, "<row r=\"{row_index}\">").expect("writing to String cannot fail");
        for (physical_index, value) in values {
            let reference = format!("{}{row_index}", column_name(physical_index + 1));
            appended.push_str(&render_cell_xml(&reference, value));
        }
        appended.push_str("</row>");
    }
    let expanded = format!(
        "{}{}{}",
        &xml[..sheet_data_end],
        appended,
        &xml[sheet_data_end..]
    );
    let next = u32::try_from(next_row + rows.len()).unwrap_or(u32::MAX);
    Ok((update_worksheet_dimension(&expanded), next))
}

/// Expands empty self-closing `<sheetData…/>` into an open/close pair.
///
/// # Errors
///
/// Returns [`ExcelError::Format`] when the worksheet has no `sheetData` element.
fn expand_self_closing_sheet_data(xml: &str) -> Result<String> {
    if xml.contains("</sheetData>") {
        return Ok(xml.to_owned());
    }
    let Some(start) = xml.find("<sheetData") else {
        return Err(ExcelError::Format(
            "worksheet does not contain sheetData".to_owned(),
        ));
    };
    let after = &xml[start..];
    let Some(rel_end) = after.find("/>") else {
        return Err(ExcelError::Format(
            "worksheet does not contain sheetData".to_owned(),
        ));
    };
    // Refuse to rewrite when `/>` belongs to a later sibling (malformed / unexpected).
    if after[..rel_end].contains('>') {
        return Err(ExcelError::Format(
            "worksheet does not contain sheetData".to_owned(),
        ));
    }
    let end = start + rel_end;
    let open_tag = &xml[start..end];
    Ok(format!(
        "{}{}></sheetData>{}",
        &xml[..start],
        open_tag,
        &xml[end + 2..]
    ))
}

fn render_cell_xml(reference: &str, value: &CellValue) -> String {
    let start = format!("<c r=\"{reference}\">");
    match value {
        CellValue::Empty | CellValue::Image(_) => format!("{start}</c>"),
        CellValue::String(text)
        | CellValue::Error(text)
        | CellValue::Hyperlink { text, .. } => format!(
            "<c r=\"{reference}\" t=\"inlineStr\"><is><t>{}</t></is></c>",
            escape_xml(text)
        ),
        CellValue::RichText(rich) => format!(
            "<c r=\"{reference}\" t=\"inlineStr\"><is><t>{}</t></is></c>",
            escape_xml(rich.text_string())
        ),
        CellValue::Bool(flag) => {
            format!("<c r=\"{reference}\" t=\"b\"><v>{}</v></c>", u8::from(*flag))
        }
        CellValue::Int(number) => format!("<c r=\"{reference}\"><v>{number}</v></c>"),
        CellValue::Float(number) => format!("<c r=\"{reference}\"><v>{number}</v></c>"),
        CellValue::Decimal(number) => format!("<c r=\"{reference}\"><v>{number}</v></c>"),
        CellValue::Date(date) => format!(
            "<c r=\"{reference}\" t=\"d\"><v>{}</v></c>",
            date.format("%Y-%m-%d")
        ),
        CellValue::DateTime(datetime) => format!(
            "<c r=\"{reference}\" t=\"d\"><v>{}</v></c>",
            datetime.format("%Y-%m-%dT%H:%M:%S")
        ),
        CellValue::Formula(formula) => {
            format!("<c r=\"{reference}\"><f>{}</f></c>", escape_xml(formula))
        }
        CellValue::Comment { value, .. } | CellValue::Images { value, .. } => {
            render_cell_xml(reference, value)
        }
    }
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

fn row_index(tag: &str) -> Option<usize> {
    attribute_value(tag, "r")?.parse().ok()
}

fn update_worksheet_dimension(xml: &str) -> String {
    let mut last_col = 1usize;
    let mut last_row = 1usize;
    let mut offset = 0;
    while let Some(relative) = xml[offset..].find("<c ") {
        let start = offset + relative;
        let Some(tag_end) = xml[start..].find('>') else {
            break;
        };
        let tag = &xml[start..start + tag_end];
        if let Some(reference) = attribute_value(tag, "r")
            && let Some((column, row)) = parse_cell_reference(reference)
        {
            last_col = last_col.max(column);
            last_row = last_row.max(row);
        }
        offset = start + tag_end + 1;
    }
    let reference = format!("{}{}:{}{}", column_name(1), 1, column_name(last_col), last_row);
    if let Some(current) = attribute_value_in_tag(xml, "dimension", "ref") {
        return xml.replacen(
            &format!(" ref=\"{current}\""),
            &format!(" ref=\"{reference}\""),
            1,
        );
    }
    xml.to_owned()
}

fn parse_cell_reference(reference: &str) -> Option<(usize, usize)> {
    let split = reference
        .find(|character: char| character.is_ascii_digit())
        .unwrap_or(reference.len());
    let (letters, digits) = reference.split_at(split);
    let mut column = 0usize;
    for character in letters.chars() {
        if !character.is_ascii_alphabetic() {
            return None;
        }
        column = column * 26 + usize::from(character.to_ascii_uppercase() as u8 - b'A') + 1;
    }
    let row = digits.parse().ok()?;
    Some((column, row))
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

fn attribute_value<'a>(xml: &'a str, attribute: &str) -> Option<&'a str> {
    let marker = format!(" {attribute}=\"");
    let start = xml.find(&marker)? + marker.len();
    let end = start + xml[start..].find('"')?;
    Some(&xml[start..end])
}

fn attribute_value_in_tag<'a>(xml: &'a str, tag: &str, attribute: &str) -> Option<&'a str> {
    let start = xml.find(&format!("<{tag}"))?;
    let end = start + xml[start..].find('>')?;
    attribute_value(&xml[start..=end], attribute)
}

fn xml_elements<'a>(xml: &'a str, tag: &'a str) -> impl Iterator<Item = &'a str> + 'a {
    let open = format!("<{tag}");
    let mut offset = 0;
    std::iter::from_fn(move || {
        let relative = xml[offset..].find(&open)?;
        let start = offset + relative;
        let end = start + xml[start..].find('>')? + 1;
        offset = end;
        Some(&xml[start..end])
    })
}

fn normalize_workbook_target(target: &str) -> Result<String> {
    let trimmed = target.trim_start_matches('/');
    if trimmed.starts_with("xl/") {
        Ok(trimmed.to_owned())
    } else {
        Ok(format!("xl/{trimmed}"))
    }
}

/// Minimal empty worksheet part used when creating a sheet absent from the template.
///
/// Prefer an open/close `sheetData` pair so [`append_sparse_rows_to_xml`] can
/// splice rows; self-closing `<sheetData/>` is still accepted and expanded.
const EMPTY_WORKSHEET_XML: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    r#"<dimension ref="A1"/><sheetData></sheetData></worksheet>"#
);

/// Builds a blank worksheet, optionally inheriting `sheetFormatPr` / `cols` from
/// the first template sheet (workbook `styles.xml` remains shared and untouched).
fn blank_worksheet_with_inherited_format(entries: &[TemplateZipEntry]) -> Vec<u8> {
    let Some(source) = entries.iter().find(|entry| {
        let lower = entry.name.to_ascii_lowercase();
        lower.starts_with("xl/worksheets/sheet") && lower.ends_with(".xml")
    }) else {
        return EMPTY_WORKSHEET_XML.as_bytes().to_vec();
    };
    let Ok(xml) = std::str::from_utf8(&source.bytes) else {
        return EMPTY_WORKSHEET_XML.as_bytes().to_vec();
    };
    let format = extract_xml_element(xml, "sheetFormatPr").unwrap_or_default();
    let cols = extract_xml_element(xml, "cols").unwrap_or_default();
    if format.is_empty() && cols.is_empty() {
        return EMPTY_WORKSHEET_XML.as_bytes().to_vec();
    }
    format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<dimension ref="A1"/>{format}{cols}<sheetData></sheetData></worksheet>"#
        ),
        format = format,
        cols = cols
    )
    .into_bytes()
}

/// Returns the first XML element named `tag`, including a self-closing form.
fn extract_xml_element(xml: &str, tag: &str) -> Option<String> {
    let open = format!("<{tag}");
    let start = xml.find(&open)?;
    let rest = &xml[start..];
    let close = format!("</{tag}>");
    if let Some(close_at) = rest.find(&close) {
        return Some(rest[..=close_at + close.len() - 1].to_owned());
    }
    let self_close = rest.find("/>")?;
    if rest[..self_close].contains('>') {
        return None;
    }
    Some(rest[..=self_close + 1].to_owned())
}

/// Picks the next unused `xl/worksheets/sheetN.xml` part name.
fn next_worksheet_part_name(entries: &[TemplateZipEntry]) -> String {
    let mut maximum = 0usize;
    for entry in entries {
        let lower = entry.name.to_ascii_lowercase();
        let Some(rest) = lower.strip_prefix("xl/worksheets/sheet") else {
            continue;
        };
        let digits: String = rest
            .chars()
            .take_while(|character| character.is_ascii_digit())
            .collect();
        if let Ok(index) = digits.parse::<usize>() {
            maximum = maximum.max(index);
        }
    }
    format!("xl/worksheets/sheet{}.xml", maximum.saturating_add(1))
}

/// Allocates the next `rIdN` relationship identifier.
fn next_relationship_id(rels_xml: &str) -> String {
    let mut maximum = 0usize;
    let mut offset = 0;
    while let Some(relative) = rels_xml[offset..].find("Id=\"rId") {
        let start = offset + relative + "Id=\"rId".len();
        let digits: String = rels_xml[start..]
            .chars()
            .take_while(|character| character.is_ascii_digit())
            .collect();
        if let Ok(index) = digits.parse::<usize>() {
            maximum = maximum.max(index);
        }
        offset = start;
    }
    format!("rId{}", maximum.saturating_add(1))
}

/// Allocates the next workbook `sheetId`.
fn next_sheet_id(workbook_xml: &str) -> usize {
    let mut maximum = 0usize;
    for element in xml_elements(workbook_xml, "sheet") {
        if let Some(value) = attribute_value(element, "sheetId")
            && let Ok(index) = value.parse::<usize>()
        {
            maximum = maximum.max(index);
        }
    }
    maximum.saturating_add(1)
}

/// Inserts `fragment` immediately before the first occurrence of `close_tag`.
fn insert_before_close_tag(xml: &str, close_tag: &str, fragment: &str) -> Result<String> {
    let Some(index) = xml.find(close_tag) else {
        return Err(ExcelError::Format(format!(
            "template XML is missing {close_tag}"
        )));
    };
    Ok(format!("{}{}{}", &xml[..index], fragment, &xml[index..]))
}

fn write_entries_to(
    writer: Box<dyn WriteSeek>,
    entries: &[TemplateZipEntry],
) -> Result<Box<dyn WriteSeek>> {
    let mut zip = ZipWriter::new(writer);
    for entry in entries {
        let mut options = SimpleFileOptions::default().compression_method(entry.compression);
        if let Some(mode) = entry.unix_mode {
            options = options.unix_permissions(mode);
        }
        if entry.is_dir {
            zip.add_directory(&entry.name, options).map_err(format_error)?;
        } else {
            zip.start_file(&entry.name, options).map_err(format_error)?;
            zip.write_all(&entry.bytes)?;
        }
    }
    zip.finish().map_err(format_error)
}

trait WriteSeek: Write + Seek {
    fn into_inner(self: Box<Self>) -> Result<Vec<u8>>;
}

impl WriteSeek for Cursor<Vec<u8>> {
    fn into_inner(self: Box<Self>) -> Result<Vec<u8>> {
        Ok((*self).into_inner())
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

fn looks_like_csv(bytes: &[u8]) -> bool {
    // XLSX / OLE are binary; treat plain printable text without ZIP/OLE magic as CSV-like.
    if looks_like_xlsx(bytes) || looks_like_xls(bytes) {
        return false;
    }
    bytes
        .iter()
        .take(64)
        .all(|byte| byte.is_ascii_whitespace() || byte.is_ascii_graphic() || *byte == b'\t')
}

fn looks_like_xlsx(bytes: &[u8]) -> bool {
    bytes.starts_with(&[0x50, 0x4B]) // ZIP / OOXML
}

fn looks_like_xls(bytes: &[u8]) -> bool {
    bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) // OLE Compound File
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_template_target_prefers_index_then_name() {
        let sheets = vec![
            TemplateSheetData {
                name: "A".to_owned(),
                cells: Vec::new(),
                next_row: 1,
            },
            TemplateSheetData {
                name: "B".to_owned(),
                cells: Vec::new(),
                next_row: 2,
            },
        ];
        // Java `.sheet(1)` / `.sheet(1, "ignored")` selects by index.
        let (index, name, create_new) = resolve_template_target(&sheets, Some(1), "ignored");
        assert_eq!((index, name.as_str(), create_new), (1, "B", false));
        // Java `.sheet("B")` selects by name when sheetNo is null.
        let (index, name, create_new) = resolve_template_target(&sheets, None, "B");
        assert_eq!((index, name.as_str(), create_new), (1, "B", false));
        // Missing name creates a new sheet after template sheets.
        let (index, name, create_new) = resolve_template_target(&sheets, None, "C");
        assert_eq!((index, name.as_str(), create_new), (2, "C", true));
    }

    #[test]
    fn validate_template_source_rejects_csv_and_xls_for_xlsx_path() {
        let csv = Path::new("demo.csv");
        let err = validate_template_source(Some(csv), None).expect_err("csv");
        assert!(err.to_string().contains("csv cannot use template"));

        let xls = Path::new("demo.xls");
        let err = validate_template_source(Some(xls), None).expect_err("xls");
        assert!(
            err.to_string().contains("legacy XLS template cannot seed an XLSX"),
            "unexpected: {err}"
        );

        let err = validate_template_source(None, Some(b"name,age\n")).expect_err("csv bytes");
        assert!(err.to_string().contains("csv cannot use template"));
    }

    #[test]
    fn append_sparse_rows_preserves_merge_cells_trailer() {
        let xml = concat!(
            "<worksheet><dimension ref=\"A1:B1\"/>",
            "<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c></row></sheetData>",
            "<mergeCells count=\"1\"><mergeCell ref=\"A1:B1\"/></mergeCells>",
            "</worksheet>"
        );
        let rows = vec![vec![(0usize, CellValue::String("appended".to_owned()))]];
        let (updated, next) = append_sparse_rows_to_xml(xml, &rows).expect("append");
        assert_eq!(next, 3);
        assert!(updated.contains("s=\"1\""));
        assert!(updated.contains("<mergeCell ref=\"A1:B1\"/>"));
        assert!(updated.contains("inlineStr"));
        assert!(updated.contains("appended"));
    }

    #[test]
    fn append_sparse_rows_expands_self_closing_sheet_data() {
        let xml = concat!(
            "<worksheet><dimension ref=\"A1\"/>",
            "<sheetData/>",
            "</worksheet>"
        );
        let rows = vec![vec![(0usize, CellValue::String("fresh".to_owned()))]];
        let (updated, next) = append_sparse_rows_to_xml(xml, &rows).expect("append");
        assert_eq!(next, 2);
        assert!(updated.contains("<sheetData><row r=\"1\">"));
        assert!(updated.contains("fresh"));
        assert!(updated.contains("</sheetData>"));
        assert!(!updated.contains("<sheetData/>"));
    }

    #[test]
    fn create_sheet_keeps_existing_styles_and_merges() {
        let template = concat!(
            "PK\x03\x04" // placeholder — build via TemplatePackage entries below
        );
        let _ = template;
        let mut package = TemplatePackage {
            entries: vec![
                TemplateZipEntry {
                    name: "[Content_Types].xml".to_owned(),
                    is_dir: false,
                    compression: CompressionMethod::Stored,
                    unix_mode: None,
                    bytes: br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_vec(),
                },
                TemplateZipEntry {
                    name: "xl/workbook.xml".to_owned(),
                    is_dir: false,
                    compression: CompressionMethod::Stored,
                    unix_mode: None,
                    bytes: br#"<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Styled" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec(),
                },
                TemplateZipEntry {
                    name: "xl/_rels/workbook.xml.rels".to_owned(),
                    is_dir: false,
                    compression: CompressionMethod::Stored,
                    unix_mode: None,
                    bytes: br#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_vec(),
                },
                TemplateZipEntry {
                    name: "xl/styles.xml".to_owned(),
                    is_dir: false,
                    compression: CompressionMethod::Stored,
                    unix_mode: None,
                    bytes: br#"<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><b/></font></fonts></styleSheet>"#.to_vec(),
                },
                TemplateZipEntry {
                    name: "xl/worksheets/sheet1.xml".to_owned(),
                    is_dir: false,
                    compression: CompressionMethod::Stored,
                    unix_mode: None,
                    bytes: br#"<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="18"/><cols><col min="1" max="1" width="20" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>"#.to_vec(),
                },
            ],
        };
        let styles_before = package
            .entries
            .iter()
            .find(|entry| entry.name == "xl/styles.xml")
            .map(|entry| entry.bytes.clone())
            .expect("styles");
        let sheet1_before = package
            .entries
            .iter()
            .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
            .map(|entry| entry.bytes.clone())
            .expect("sheet1");

        package.create_sheet("NewSheet").expect("create");
        package
            .append_rows(
                "NewSheet",
                &[vec![(0usize, CellValue::String("fresh".to_owned()))]],
            )
            .expect("append");

        let styles_after = package
            .entries
            .iter()
            .find(|entry| entry.name == "xl/styles.xml")
            .map(|entry| entry.bytes.clone())
            .expect("styles");
        let sheet1_after = package
            .entries
            .iter()
            .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
            .map(|entry| entry.bytes.clone())
            .expect("sheet1");
        assert_eq!(styles_before, styles_after, "styles.xml must be untouched");
        assert_eq!(
            sheet1_before, sheet1_after,
            "existing sheet XML (incl. mergeCells) must be untouched"
        );
        assert!(
            package
                .sheet_names()
                .expect("names")
                .iter()
                .any(|name| name == "NewSheet")
        );
        let new_sheet = package
            .entry_xml("xl/worksheets/sheet2.xml")
            .expect("new sheet xml");
        assert!(new_sheet.contains("fresh"));
        assert!(
            new_sheet.contains("defaultRowHeight=\"18\""),
            "new sheet should inherit sheetFormatPr: {new_sheet}"
        );
        assert!(
            new_sheet.contains("customWidth=\"1\""),
            "new sheet should inherit cols: {new_sheet}"
        );
        assert!(
            !new_sheet.contains("mergeCell"),
            "new sheet must not copy merges from the template sheet"
        );
    }
}
