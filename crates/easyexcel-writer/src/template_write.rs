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

use crate::MergeRange;
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
        Ok(self
            .workbook_sheets()?
            .into_iter()
            .map(|(name, _)| name)
            .collect())
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
        let selected = sheets
            .get(index)
            .ok_or_else(|| ExcelError::SheetNotFound(format!("sheet index {index}")))?;
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
        if self.sheet_names()?.iter().any(|name| name == sheet_name) {
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
        let updated_types = insert_before_close_tag(&content_types_xml, "</Types>", &override_tag)?;

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

    /// Appends typed rows into a worksheet's `sheetData`.
    ///
    /// # Errors
    ///
    /// Returns a format error when the worksheet XML cannot be updated.
    pub(crate) fn append_rows(
        &mut self,
        sheet_name: &str,
        rows: &[Vec<(usize, CellValue)>],
    ) -> Result<u32> {
        self.append_rows_with_heights(sheet_name, rows, &[])
    }

    /// Appends rows and applies optional per-row heights to the newly created
    /// row elements.
    pub(crate) fn append_rows_with_heights(
        &mut self,
        sheet_name: &str,
        rows: &[Vec<(usize, CellValue)>],
        row_heights: &[Option<u16>],
    ) -> Result<u32> {
        self.append_rows_with_layout(sheet_name, rows, row_heights, &[])
    }

    /// Appends rows with optional row heights and per-cell workbook style indexes.
    pub(crate) fn append_rows_with_layout(
        &mut self,
        sheet_name: &str,
        rows: &[Vec<(usize, CellValue)>],
        row_heights: &[Option<u16>],
        cell_styles: &[Vec<Option<u32>>],
    ) -> Result<u32> {
        self.append_rows_with_layout_and_absent(sheet_name, rows, row_heights, cell_styles, &[])
    }

    /// Appends rows while preserving Java `null` row gaps without creating
    /// empty OOXML `<row>` elements for those positions.
    pub(crate) fn append_rows_with_layout_and_absent(
        &mut self,
        sheet_name: &str,
        rows: &[Vec<(usize, CellValue)>],
        row_heights: &[Option<u16>],
        cell_styles: &[Vec<Option<u32>>],
        absent_rows: &[bool],
    ) -> Result<u32> {
        if rows.is_empty() {
            return self.next_row_for_sheet(sheet_name);
        }
        if !absent_rows.is_empty() && absent_rows.len() != rows.len() {
            return Err(ExcelError::Format(
                "template absent-row count does not match appended row count".to_owned(),
            ));
        }
        if !row_heights.is_empty() && row_heights.len() != rows.len() {
            return Err(ExcelError::Format(
                "template row-height count does not match appended row count".to_owned(),
            ));
        }
        if !cell_styles.is_empty()
            && (cell_styles.len() != rows.len()
                || cell_styles
                    .iter()
                    .zip(rows)
                    .any(|(styles, row)| styles.len() != row.len()))
        {
            return Err(ExcelError::Format(
                "template cell-style shape does not match appended rows".to_owned(),
            ));
        }
        let path = self.worksheet_path_by_name(sheet_name)?;
        let entry = self
            .entries
            .iter_mut()
            .find(|entry| entry.name.eq_ignore_ascii_case(&path))
            .ok_or_else(|| ExcelError::Format(format!("template does not contain {path}")))?;
        let xml = String::from_utf8(std::mem::take(&mut entry.bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let (updated, next_row) =
            append_sparse_rows_to_xml(&xml, rows, row_heights, cell_styles, absent_rows)?;
        entry.bytes = updated.into_bytes();
        Ok(next_row)
    }

    /// Applies column widths and absolute merged regions to one preserved
    /// worksheet part.
    ///
    /// This is the OOXML equivalent of Java annotation-generated
    /// `AbstractHeadColumnWidthStyleStrategy` and
    /// `OnceAbsoluteMergeStrategy` callbacks. Existing package entries and
    /// style indexes remain untouched.
    pub(crate) fn apply_sheet_layout(
        &mut self,
        sheet_name: &str,
        column_widths: &[(u16, u16)],
        merge_ranges: &[MergeRange],
    ) -> Result<()> {
        if column_widths.is_empty() && merge_ranges.is_empty() {
            return Ok(());
        }
        let path = self.worksheet_path_by_name(sheet_name)?;
        let entry = self
            .entries
            .iter_mut()
            .find(|entry| entry.name.eq_ignore_ascii_case(&path))
            .ok_or_else(|| ExcelError::Format(format!("template does not contain {path}")))?;
        let mut xml = String::from_utf8(std::mem::take(&mut entry.bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        if !column_widths.is_empty() {
            xml = apply_column_widths_to_xml(&xml, column_widths)?;
        }
        if !merge_ranges.is_empty() {
            xml = apply_merge_ranges_to_xml(&xml, merge_ranges)?;
        }
        entry.bytes = xml.into_bytes();
        Ok(())
    }

    /// Imports styles compiled by `rust_xlsxwriter` into the preserved
    /// template style table and returns the destination style index for each
    /// compiler worksheet row.
    pub(crate) fn import_compiled_styles(
        &mut self,
        compiled_xlsx: &[u8],
        style_count: usize,
    ) -> Result<Vec<u32>> {
        if style_count == 0 {
            return Ok(Vec::new());
        }
        let compiled = Self::from_bytes(compiled_xlsx)?;
        let source_styles = compiled.entry_xml("xl/styles.xml")?;
        let (_, source_sheet_path) = compiled.worksheet_path_by_index(0)?;
        let source_sheet = compiled.entry_xml(&source_sheet_path)?;
        let source_indexes = (1..=style_count)
            .map(|row| {
                cell_style_index(&source_sheet, &format!("A{row}")).ok_or_else(|| {
                    ExcelError::Format(format!(
                        "compiled template style cell A{row} has no style index"
                    ))
                })
            })
            .collect::<Result<Vec<_>>>()?;

        let destination = self
            .entries
            .iter_mut()
            .find(|entry| entry.name.eq_ignore_ascii_case("xl/styles.xml"))
            .ok_or_else(|| ExcelError::Format("template missing xl/styles.xml".to_owned()))?;
        let destination_styles = String::from_utf8(std::mem::take(&mut destination.bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let (updated, mapped) =
            merge_compiled_styles(&destination_styles, &source_styles, &source_indexes)?;
        destination.bytes = updated.into_bytes();
        Ok(mapped)
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
        String::from_utf8(entry.bytes.clone())
            .map_err(|error| ExcelError::Format(error.to_string()))
    }
}

/// Returns whether [`crate::WriteOptions`] carries a template source.
///
/// Corresponds to Java `WriteWorkbook.templateFile` / `templateInputStream`
/// being non-null.
#[must_use]
pub(crate) fn has_template(template_file: Option<&Path>, template_bytes: Option<&[u8]>) -> bool {
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
    let mut workbook: Xlsx<_> = open_workbook_from_rs(Cursor::new(bytes)).map_err(|error| {
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
    row_heights: &[Option<u16>],
    cell_styles: &[Vec<Option<u32>>],
    absent_rows: &[bool],
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
        if absent_rows.get(row_offset).copied().unwrap_or(false) {
            continue;
        }
        if let Some(height) = row_heights.get(row_offset).copied().flatten() {
            write!(
                appended,
                "<row r=\"{row_index}\" ht=\"{height}\" customHeight=\"1\">"
            )
            .expect("writing to String cannot fail");
        } else {
            write!(appended, "<row r=\"{row_index}\">").expect("writing to String cannot fail");
        }
        for (cell_offset, (physical_index, value)) in values.iter().enumerate() {
            let reference = format!("{}{row_index}", column_name(physical_index + 1));
            let style = cell_styles
                .get(row_offset)
                .and_then(|styles| styles.get(cell_offset))
                .copied()
                .flatten();
            appended.push_str(&render_cell_xml(&reference, value, style));
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

fn apply_column_widths_to_xml(xml: &str, widths: &[(u16, u16)]) -> Result<String> {
    let mut tags = String::new();
    for (column, width) in widths {
        let one_based = u32::from(*column) + 1;
        write!(
            tags,
            "<col min=\"{one_based}\" max=\"{one_based}\" width=\"{width}\" customWidth=\"1\"/>"
        )
        .expect("writing to String cannot fail");
    }
    if let Some(end) = xml.find("</cols>") {
        return Ok(format!("{}{}{}", &xml[..end], tags, &xml[end..]));
    }
    if let Some(start) = xml.find("<cols") {
        let Some(relative_end) = xml[start..].find("/>") else {
            return Err(ExcelError::Format(
                "worksheet contains malformed cols element".to_owned(),
            ));
        };
        let end = start + relative_end + 2;
        return Ok(format!(
            "{}<cols>{}</cols>{}",
            &xml[..start],
            tags,
            &xml[end..]
        ));
    }
    let insertion = xml
        .find("<sheetData")
        .ok_or_else(|| ExcelError::Format("worksheet does not contain sheetData".to_owned()))?;
    Ok(format!(
        "{}<cols>{}</cols>{}",
        &xml[..insertion],
        tags,
        &xml[insertion..]
    ))
}

fn apply_merge_ranges_to_xml(xml: &str, ranges: &[MergeRange]) -> Result<String> {
    let mut refs = Vec::new();
    for range in ranges {
        let reference = format!(
            "{}{}:{}{}",
            column_name(usize::from(range.first_column) + 1),
            range.first_row + 1,
            column_name(usize::from(range.last_column) + 1),
            range.last_row + 1
        );
        if !xml.contains(&format!("ref=\"{reference}\"")) {
            refs.push(reference);
        }
    }
    if refs.is_empty() {
        return Ok(xml.to_owned());
    }
    let tags = refs
        .iter()
        .map(|reference| format!("<mergeCell ref=\"{reference}\"/>"))
        .collect::<String>();
    if let Some(start) = xml.find("<mergeCells") {
        let tag_end = start
            + xml[start..]
                .find('>')
                .ok_or_else(|| ExcelError::Format("malformed mergeCells element".to_owned()))?;
        let close = xml[tag_end + 1..]
            .find("</mergeCells>")
            .map(|offset| tag_end + 1 + offset)
            .ok_or_else(|| ExcelError::Format("malformed mergeCells element".to_owned()))?;
        let current_count = attribute_value(&xml[start..=tag_end], "count")
            .and_then(|count| count.parse::<usize>().ok())
            .unwrap_or_else(|| xml[tag_end + 1..close].matches("<mergeCell").count());
        let new_count = current_count.saturating_add(refs.len());
        let mut updated = xml.to_owned();
        if let Some(count) = attribute_value(&xml[start..=tag_end], "count") {
            updated = updated.replacen(
                &format!(" count=\"{count}\""),
                &format!(" count=\"{new_count}\""),
                1,
            );
        }
        let close = updated
            .find("</mergeCells>")
            .ok_or_else(|| ExcelError::Format("malformed mergeCells element".to_owned()))?;
        return Ok(format!(
            "{}{}{}",
            &updated[..close],
            tags,
            &updated[close..]
        ));
    }
    let insertion = xml
        .find("</sheetData>")
        .map(|index| index + "</sheetData>".len())
        .ok_or_else(|| ExcelError::Format("worksheet does not contain sheetData".to_owned()))?;
    Ok(format!(
        "{}<mergeCells count=\"{}\">{}</mergeCells>{}",
        &xml[..insertion],
        refs.len(),
        tags,
        &xml[insertion..]
    ))
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

fn render_cell_xml(reference: &str, value: &CellValue, style: Option<u32>) -> String {
    let style_attribute = style
        .map(|index| format!(" s=\"{index}\""))
        .unwrap_or_default();
    let start = format!("<c r=\"{reference}\"{style_attribute}>");
    match value {
        CellValue::Empty | CellValue::Image(_) => format!("{start}</c>"),
        CellValue::String(text) | CellValue::Error(text) | CellValue::Hyperlink { text, .. } => {
            format!(
                "<c r=\"{reference}\"{style_attribute} t=\"inlineStr\"><is><t>{}</t></is></c>",
                escape_xml(text)
            )
        }
        CellValue::RichText(rich) => format!(
            "<c r=\"{reference}\"{style_attribute} t=\"inlineStr\"><is><t>{}</t></is></c>",
            escape_xml(rich.text_string())
        ),
        CellValue::Bool(flag) => {
            format!(
                "<c r=\"{reference}\"{style_attribute} t=\"b\"><v>{}</v></c>",
                u8::from(*flag)
            )
        }
        CellValue::Int(number) => format!("{start}<v>{number}</v></c>"),
        CellValue::Float(number) => format!("{start}<v>{number}</v></c>"),
        CellValue::Decimal(number) => {
            if crate::decimal_integer_requires_text(number).unwrap_or(false) {
                format!(
                    "<c r=\"{reference}\"{style_attribute} t=\"inlineStr\"><is><t>{}</t></is></c>",
                    escape_xml(&number.to_plain_string())
                )
            } else {
                format!("{start}<v>{number}</v></c>")
            }
        }
        CellValue::Date(date) => format!(
            "<c r=\"{reference}\"{style_attribute} t=\"d\"><v>{}</v></c>",
            date.format("%Y-%m-%d")
        ),
        CellValue::DateTime(datetime) => format!(
            "<c r=\"{reference}\"{style_attribute} t=\"d\"><v>{}</v></c>",
            datetime.format("%Y-%m-%dT%H:%M:%S")
        ),
        CellValue::Formula(formula) => {
            format!("{start}<f>{}</f></c>", escape_xml(formula))
        }
        CellValue::Comment { value, .. } | CellValue::Images { value, .. } => {
            render_cell_xml(reference, value, style)
        }
    }
}

fn cell_style_index(sheet_xml: &str, reference: &str) -> Option<usize> {
    let marker = format!("<c r=\"{reference}\"");
    let (_, cell) = sheet_xml.split_once(&marker)?;
    let tag = cell.split_once('>')?.0;
    attribute_value(tag, "s")?.parse().ok()
}

fn merge_compiled_styles(
    destination: &str,
    source: &str,
    source_indexes: &[usize],
) -> Result<(String, Vec<u32>)> {
    let source_fonts = collection_elements(source, "fonts", "font")?;
    let source_fills = collection_elements(source, "fills", "fill")?;
    let source_borders = collection_elements(source, "borders", "border")?;
    let source_xfs = collection_elements(source, "cellXfs", "xf")?;
    let (mut updated, font_indexes) =
        merge_component_collection(destination, "fonts", "font", &source_fonts)?;
    let (next, fill_indexes) =
        merge_component_collection(&updated, "fills", "fill", &source_fills)?;
    updated = next;
    let (next, border_indexes) =
        merge_component_collection(&updated, "borders", "border", &source_borders)?;
    updated = next;

    let mut imported = std::collections::HashMap::new();
    let mut appended_xfs = Vec::new();
    let destination_xfs = collection_elements(&updated, "cellXfs", "xf")?;
    let mut mapped = Vec::with_capacity(source_indexes.len());
    for source_index in source_indexes {
        if let Some(destination_index) = imported.get(source_index).copied() {
            mapped.push(destination_index);
            continue;
        }
        let source_xf = source_xfs.get(*source_index).ok_or_else(|| {
            ExcelError::Format(format!(
                "compiled style index {source_index} is out of range"
            ))
        })?;
        let mut xf = source_xf.clone();
        remap_index_attribute(&mut xf, "fontId", &font_indexes)?;
        remap_index_attribute(&mut xf, "fillId", &fill_indexes)?;
        remap_index_attribute(&mut xf, "borderId", &border_indexes)?;
        remap_number_format(&mut updated, source, &mut xf)?;
        let destination_index = destination_xfs
            .iter()
            .chain(appended_xfs.iter())
            .position(|existing| existing == &xf)
            .map(|index| u32::try_from(index).unwrap_or(u32::MAX))
            .unwrap_or_else(|| {
                let index =
                    u32::try_from(destination_xfs.len() + appended_xfs.len()).unwrap_or(u32::MAX);
                appended_xfs.push(xf);
                index
            });
        if destination_index == u32::MAX {
            return Err(ExcelError::Format(
                "template cell style index overflow".to_owned(),
            ));
        }
        imported.insert(*source_index, destination_index);
        mapped.push(destination_index);
    }
    updated = append_collection(&updated, "cellXfs", "xf", &appended_xfs)?;
    Ok((updated, mapped))
}

fn remap_index_attribute(xml: &mut String, name: &str, indexes: &[usize]) -> Result<()> {
    let Some(value) = attribute_value(xml, name) else {
        return Ok(());
    };
    let index = value
        .parse::<usize>()
        .map_err(|_| ExcelError::Format(format!("invalid {name} in compiled style")))?;
    let mapped = indexes.get(index).ok_or_else(|| {
        ExcelError::Format(format!(
            "compiled style {name} index {index} is out of range"
        ))
    })?;
    replace_attribute(xml, name, &mapped.to_string())
}

fn merge_component_collection(
    xml: &str,
    collection: &str,
    child: &str,
    source: &[String],
) -> Result<(String, Vec<usize>)> {
    let destination = collection_elements(xml, collection, child)?;
    let mut appended = Vec::new();
    let mut indexes = Vec::with_capacity(source.len());
    for component in source {
        let index = destination
            .iter()
            .chain(appended.iter())
            .position(|existing| existing == component)
            .unwrap_or_else(|| {
                let index = destination.len() + appended.len();
                appended.push(component.clone());
                index
            });
        indexes.push(index);
    }
    Ok((
        append_collection(xml, collection, child, &appended)?,
        indexes,
    ))
}

fn remap_number_format(destination: &mut String, source: &str, xf: &mut String) -> Result<()> {
    let Some(value) = attribute_value(xf, "numFmtId") else {
        return Ok(());
    };
    let source_id = value
        .parse::<u32>()
        .map_err(|_| ExcelError::Format("invalid numFmtId in compiled style".to_owned()))?;
    if source_id < 164 {
        return Ok(());
    }
    let source_formats = optional_collection_elements(source, "numFmts", "numFmt")?;
    let source_format = source_formats
        .iter()
        .find(|format| attribute_value(format, "numFmtId") == Some(value))
        .ok_or_else(|| {
            ExcelError::Format(format!("compiled style is missing numFmtId {source_id}"))
        })?;
    let code = attribute_value(source_format, "formatCode")
        .ok_or_else(|| ExcelError::Format("compiled numFmt has no formatCode".to_owned()))?;
    let destination_formats = optional_collection_elements(destination, "numFmts", "numFmt")?;
    if let Some(existing) = destination_formats
        .iter()
        .find(|format| attribute_value(format, "formatCode") == Some(code))
    {
        let id = attribute_value(existing, "numFmtId")
            .ok_or_else(|| ExcelError::Format("template numFmt has no id".to_owned()))?;
        return replace_attribute(xf, "numFmtId", id);
    }
    let next_id = destination_formats
        .iter()
        .filter_map(|format| attribute_value(format, "numFmtId")?.parse::<u32>().ok())
        .max()
        .unwrap_or(163)
        .saturating_add(1)
        .max(164);
    let mut imported = source_format.clone();
    replace_attribute(&mut imported, "numFmtId", &next_id.to_string())?;
    *destination =
        append_optional_collection(destination, "numFmts", "numFmt", &[imported], "<fonts")?;
    replace_attribute(xf, "numFmtId", &next_id.to_string())
}

fn collection_elements(xml: &str, collection: &str, child: &str) -> Result<Vec<String>> {
    let (inner, _) = collection_inner(xml, collection)?
        .ok_or_else(|| ExcelError::Format(format!("styles.xml is missing {collection}")))?;
    Ok(extract_elements(inner, child))
}

fn optional_collection_elements(xml: &str, collection: &str, child: &str) -> Result<Vec<String>> {
    Ok(collection_inner(xml, collection)?
        .map(|(inner, _)| extract_elements(inner, child))
        .unwrap_or_default())
}

fn collection_inner<'a>(
    xml: &'a str,
    collection: &str,
) -> Result<Option<(&'a str, (usize, usize, usize))>> {
    let marker = format!("<{collection}");
    let Some(start) = xml.find(&marker) else {
        return Ok(None);
    };
    let open_end = start
        + xml[start..]
            .find('>')
            .ok_or_else(|| ExcelError::Format(format!("malformed {collection} element")))?;
    let close_marker = format!("</{collection}>");
    let close = open_end
        + 1
        + xml[open_end + 1..]
            .find(&close_marker)
            .ok_or_else(|| ExcelError::Format(format!("malformed {collection} element")))?;
    Ok(Some((&xml[open_end + 1..close], (start, open_end, close))))
}

fn extract_elements(xml: &str, child: &str) -> Vec<String> {
    let marker = format!("<{child}");
    let close_marker = format!("</{child}>");
    let mut elements = Vec::new();
    let mut offset = 0;
    while let Some(relative_start) = xml[offset..].find(&marker) {
        let start = offset + relative_start;
        let Some(relative_open_end) = xml[start..].find('>') else {
            break;
        };
        let open_end = start + relative_open_end;
        let end = if xml[..=open_end].ends_with("/>") {
            open_end + 1
        } else if let Some(relative_close) = xml[open_end + 1..].find(&close_marker) {
            open_end + 1 + relative_close + close_marker.len()
        } else {
            break;
        };
        elements.push(xml[start..end].to_owned());
        offset = end;
    }
    elements
}

fn append_collection(
    xml: &str,
    collection: &str,
    child: &str,
    elements: &[String],
) -> Result<String> {
    if elements.is_empty() {
        return Ok(xml.to_owned());
    }
    let (_, (start, open_end, close)) = collection_inner(xml, collection)?
        .ok_or_else(|| ExcelError::Format(format!("styles.xml is missing {collection}")))?;
    let current = extract_elements(&xml[open_end + 1..close], child).len();
    let mut opening = xml[start..=open_end].to_owned();
    set_count_attribute(&mut opening, current + elements.len())?;
    Ok(format!(
        "{}{}{}{}{}",
        &xml[..start],
        opening,
        &xml[open_end + 1..close],
        elements.concat(),
        &xml[close..]
    ))
}

fn append_optional_collection(
    xml: &str,
    collection: &str,
    child: &str,
    elements: &[String],
    before: &str,
) -> Result<String> {
    if collection_inner(xml, collection)?.is_some() {
        return append_collection(xml, collection, child, elements);
    }
    let insertion = xml
        .find(before)
        .ok_or_else(|| ExcelError::Format(format!("styles.xml is missing {before}")))?;
    Ok(format!(
        "{}<{} count=\"{}\">{}</{}>{}",
        &xml[..insertion],
        collection,
        elements.len(),
        elements.concat(),
        collection,
        &xml[insertion..]
    ))
}

fn set_count_attribute(opening: &mut String, count: usize) -> Result<()> {
    if attribute_value(opening, "count").is_some() {
        replace_attribute(opening, "count", &count.to_string())
    } else {
        let insertion = opening
            .find('>')
            .ok_or_else(|| ExcelError::Format("malformed style collection".to_owned()))?;
        opening.insert_str(insertion, &format!(" count=\"{count}\""));
        Ok(())
    }
}

fn replace_attribute(xml: &mut String, name: &str, replacement: &str) -> Result<()> {
    let marker = format!("{name}=\"");
    let start = xml
        .find(&marker)
        .ok_or_else(|| ExcelError::Format(format!("missing {name} attribute")))?
        + marker.len();
    let end = start
        + xml[start..]
            .find('"')
            .ok_or_else(|| ExcelError::Format(format!("unterminated {name} attribute")))?;
    xml.replace_range(start..end, replacement);
    Ok(())
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
    let reference = format!(
        "{}{}:{}{}",
        column_name(1),
        1,
        column_name(last_col),
        last_row
    );
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
            zip.add_directory(&entry.name, options)
                .map_err(format_error)?;
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
            err.to_string()
                .contains("legacy XLS template cannot seed an XLSX"),
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
        let (updated, next) = append_sparse_rows_to_xml(xml, &rows, &[], &[], &[]).expect("append");
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
        let (updated, next) = append_sparse_rows_to_xml(xml, &rows, &[], &[], &[]).expect("append");
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

    #[test]
    fn importing_same_compiled_style_reuses_workbook_components_and_xf() {
        let mut template = Workbook::new();
        template
            .add_worksheet()
            .write_string_with_format(
                0,
                0,
                "seed",
                &Format::new().set_bold().set_font_color(0x0000_00ff),
            )
            .expect("template seed");
        let template_bytes = template.save_to_buffer().expect("template bytes");
        let mut package = TemplatePackage::from_bytes(&template_bytes).expect("template package");

        let mut compiler = Workbook::new();
        compiler
            .add_worksheet()
            .write_blank(
                0,
                0,
                &Format::new()
                    .set_italic()
                    .set_font_color(0x00ff_0000)
                    .set_num_format("0.000"),
            )
            .expect("compiled style");
        let compiled = compiler.save_to_buffer().expect("compiled bytes");

        let first = package
            .import_compiled_styles(&compiled, 1)
            .expect("first import");
        let after_first = package.entry_xml("xl/styles.xml").expect("styles");
        let second = package
            .import_compiled_styles(&compiled, 1)
            .expect("second import");
        let after_second = package.entry_xml("xl/styles.xml").expect("styles");

        assert_eq!(first, second);
        assert_eq!(after_first, after_second);
    }
}
