//! OOXML-preserving XLSX template filling.

use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::path::Path;

use easyexcel_core::{ExcelError, Result};
use zip::CompressionMethod;
use zip::read::ZipArchive;
use zip::write::{SimpleFileOptions, ZipWriter};

/// Value accepted by [`TemplateData`] placeholder insertion methods.
pub trait IntoTemplateValue {
    /// Converts the value to its scalar template representation.
    fn into_template_value(self) -> String;
}

impl<T> IntoTemplateValue for T
where
    T: std::fmt::Display,
{
    fn into_template_value(self) -> String {
        self.to_string()
    }
}

/// Scalar values used to replace `{key}` placeholders in OOXML text nodes.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TemplateData {
    values: BTreeMap<String, String>,
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
#[derive(Debug, Clone, Default, PartialEq, Eq)]
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
    ) -> Option<String> {
        self.values.insert(key.into(), value.into_template_value())
    }

    /// Returns all values in deterministic key order.
    #[must_use]
    pub const fn values(&self) -> &BTreeMap<String, String> {
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

trait ReadSeek: Read + Seek {}

impl<T: Read + Seek> ReadSeek for T {}

trait WriteSeek: Write + Seek {}

impl<T: Write + Seek> WriteSeek for T {}

type ArchiveWriter = ZipWriter<Box<dyn WriteSeek>>;

/// Fills scalar `{key}` placeholders while preserving the XLSX package structure.
///
/// The template and output paths may be identical because the source archive is
/// fully loaded before the destination is opened.
///
/// # Errors
///
/// Returns an I/O or format error for invalid ZIP/OOXML input or output failures.
pub fn fill_xlsx_template(template: &Path, output: &Path, data: &TemplateData) -> Result<()> {
    let mut entries = load_entries(template)?;
    replace_xml_placeholders(&mut entries, data)?;
    write_entries(output, &entries)
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
    let mut entries = load_entries(template)?;
    replace_collection_placeholders(&mut entries, data, config);
    write_entries(output, &entries)
}

fn replace_collection_placeholders(
    entries: &mut [TemplateEntry],
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
        entry.name.starts_with("xl/worksheets/")
            && Path::new(&entry.name)
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
            entry.bytes = expanded.into_bytes();
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
    let mut rows = String::new();
    for (offset, data) in wrapper.rows().iter().enumerate() {
        let filled = fill_row_cells(
            row,
            data,
            wrapper.name(),
            shared_strings,
            config.get_auto_style(),
        );
        rows.push_str(&shift_row(&filled, offset, 0));
    }
    let shift = if config.get_force_new_row() {
        wrapper.rows().len().saturating_sub(1)
    } else {
        0
    };
    let suffix = shift_rows(&xml[end..], shift);
    Some(format!("{}{}{}", &xml[..start], rows, suffix))
}

fn expand_horizontal_cells(
    xml: &str,
    wrapper: &FillWrapper,
    shared_strings: &[String],
) -> Option<String> {
    let (row_start, row_end, row, cell_start, cell_end, cell) =
        find_collection_row(xml, wrapper, shared_strings)?;
    let mut cells = String::new();
    for (offset, data) in wrapper.rows().iter().enumerate() {
        let filled = fill_cell(cell, data, wrapper.name(), shared_strings, true);
        cells.push_str(&shift_row(&filled, 0, offset));
    }
    let expanded_row = format!("{}{}{}", &row[..cell_start], cells, &row[cell_end..]);
    Some(format!(
        "{}{}{}",
        &xml[..row_start],
        expanded_row,
        &xml[row_end..]
    ))
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
    while let Some(relative_start) = row[offset..].find("<c") {
        let start = offset + relative_start;
        let end = start + row[start..].find("</c>")? + 4;
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
    while let Some(relative_start) = row[offset..].find("<c") {
        let start = offset + relative_start;
        let Some(relative_end) = row[start..].find("</c>") else {
            break;
        };
        let end = start + relative_end + 4;
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
    value.contains(&format!("{{{prefix}"))
}

fn replace_collection_values(value: &str, data: &TemplateData, prefix: Option<&str>) -> String {
    data.values()
        .iter()
        .fold(value.to_owned(), |result, (key, value)| {
            let marker =
                prefix.map_or_else(|| format!("{{.{key}}}"), |name| format!("{{{name}.{key}}}"));
            result.replace(&marker, value)
        })
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
    load_entries_from(Box::new(File::open(path)?))
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

fn replace_xml_placeholders(entries: &mut [TemplateEntry], data: &TemplateData) -> Result<()> {
    for entry in entries.iter_mut().filter(|entry| {
        !entry.is_dir
            && Path::new(&entry.name)
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("xml"))
    }) {
        let xml = String::from_utf8(std::mem::take(&mut entry.bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        entry.bytes = replace_placeholders(&xml, data.values()).into_bytes();
    }
    Ok(())
}

fn replace_placeholders(xml: &str, values: &BTreeMap<String, String>) -> String {
    values.iter().fold(xml.to_owned(), |content, (key, value)| {
        content.replace(&format!("{{{key}}}"), &escape_xml(value))
    })
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

fn write_entries(path: &Path, entries: &[TemplateEntry]) -> Result<()> {
    match File::create(path) {
        Ok(writer) => write_file_entries(writer, entries),
        Err(error) => Err(error.into()),
    }
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
