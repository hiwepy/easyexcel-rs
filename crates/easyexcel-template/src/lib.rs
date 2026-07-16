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
