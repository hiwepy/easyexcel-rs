use std::collections::HashMap;
use std::io::{BufRead, BufReader, Read, Seek};

use easyexcel_core::{ExcelError, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Decoder, Reader as XmlReader, XmlVersion};
use zip::ZipArchive;

const MAX_XLSX_ROW_NUMBER: u32 = 1_048_576;

type Relationships = HashMap<String, (String, String)>;

pub(crate) struct XlsxRowMetadata<R> {
    archive: ZipArchive<R>,
    path_cache: HashMap<String, String>,
    sheet_paths: HashMap<String, String>,
}

impl<R: Read + Seek> XlsxRowMetadata<R> {
    pub(crate) fn new(input: R) -> Result<Self> {
        let mut archive = ZipArchive::new(input).map_err(format_error)?;
        let path_cache = path_cache(&archive);
        let package_relationships = read_relationships(&mut archive, &path_cache, "_rels/.rels")?;
        let workbook_target = package_relationships
            .values()
            .find(|(_, relationship_type)| relationship_type.ends_with("/officeDocument"))
            .map(|(target, _)| target)
            .ok_or_else(|| {
                ExcelError::Format("officeDocument relationship not found".to_owned())
            })?;
        let workbook_path = resolve_target("", workbook_target)?;
        let workbook_relationships_path = relationship_part_name(&workbook_path);
        let workbook_relationships =
            read_relationships(&mut archive, &path_cache, &workbook_relationships_path)?;
        let sheet_paths = read_sheet_paths(
            &mut archive,
            &path_cache,
            &workbook_path,
            &workbook_relationships,
        )?;
        for path in sheet_paths.values() {
            if !path_cache.contains_key(&path.to_ascii_lowercase()) {
                return Err(ExcelError::Format(format!(
                    "worksheet part not found: {path}"
                )));
            }
        }
        Ok(Self {
            archive,
            path_cache,
            sheet_paths,
        })
    }

    pub(crate) fn last_explicit_row(&mut self, sheet_name: &str) -> Result<Option<u32>> {
        let path = self
            .sheet_paths
            .get(sheet_name)
            .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_owned()))?;
        let actual_path = cached_path(&self.path_cache, path);
        let file = self.archive.by_name(actual_path).map_err(format_error)?;
        scan_last_row(BufReader::new(file))
    }
}

fn path_cache<R: Read + Seek>(archive: &ZipArchive<R>) -> HashMap<String, String> {
    let mut paths = HashMap::with_capacity(archive.len());
    for name in archive.file_names() {
        paths.insert(name.to_ascii_lowercase(), name.to_owned());
    }
    paths
}

fn cached_path<'a>(cache: &'a HashMap<String, String>, path: &'a str) -> &'a str {
    cache
        .get(&path.to_ascii_lowercase())
        .map_or(path, String::as_str)
}

fn xml_reader<'a, R: Read + Seek>(
    archive: &'a mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    path: &str,
) -> Result<XmlReader<BufReader<zip::read::ZipFile<'a, R>>>> {
    let file = archive
        .by_name(cached_path(cache, path))
        .map_err(format_error)?;
    let mut reader = XmlReader::from_reader(BufReader::new(file));
    let config = reader.config_mut();
    config.check_end_names = false;
    config.check_comments = false;
    config.expand_empty_elements = true;
    Ok(reader)
}

fn read_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    path: &str,
) -> Result<Relationships> {
    let mut reader = xml_reader(archive, cache, path)?;
    let mut relationships = HashMap::new();
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"Relationship" => {
                let attributes = attributes(&element, reader.decoder())?;
                if attributes
                    .get("TargetMode")
                    .is_some_and(|mode| mode == "External")
                {
                    continue;
                }
                let Some(id) = attributes.get("Id") else {
                    continue;
                };
                let target = attributes.get("Target").cloned().unwrap_or_default();
                let relationship_type = attributes.get("Type").cloned().unwrap_or_default();
                relationships.insert(id.clone(), (target, relationship_type));
            }
            Event::End(element) if element.local_name().as_ref() == b"Relationships" => break,
            Event::Eof => {
                return Err(ExcelError::Format(format!(
                    "unexpected end of XML in {path}"
                )));
            }
            _ => {}
        }
    }
    Ok(relationships)
}

fn read_sheet_paths<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    workbook_path: &str,
    relationships: &Relationships,
) -> Result<HashMap<String, String>> {
    let mut reader = xml_reader(archive, cache, workbook_path)?;
    let mut sheets = HashMap::new();
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"sheet" => {
                let sheet_attributes = attributes(&element, reader.decoder())?;
                let name = sheet_attributes
                    .get("name")
                    .ok_or_else(|| ExcelError::Format("sheet name is missing".to_owned()))?;
                let relationship_id = sheet_attributes.get("id").ok_or_else(|| {
                    ExcelError::Format("sheet relationship is missing".to_owned())
                })?;
                let (target, relationship_type) =
                    relationships.get(relationship_id).ok_or_else(|| {
                        ExcelError::Format(format!(
                            "sheet relationship not found: {relationship_id}"
                        ))
                    })?;
                if relationship_type.ends_with("/worksheet") {
                    sheets.insert(name.clone(), resolve_target(workbook_path, target)?);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"workbook" => break,
            Event::Eof => {
                return Err(ExcelError::Format(
                    "unexpected end of XML in workbook".to_owned(),
                ));
            }
            _ => {}
        }
    }
    Ok(sheets)
}

fn attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<HashMap<String, String>> {
    let mut values = HashMap::new();
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(format_error)?;
        let key = std::str::from_utf8(attribute.key.local_name().as_ref())
            .map_err(format_error)?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(format_error)?
            .into_owned();
        values.insert(key, value);
    }
    Ok(values)
}

fn scan_last_row<R: BufRead>(input: R) -> Result<Option<u32>> {
    let mut reader = XmlReader::from_reader(input);
    reader.config_mut().expand_empty_elements = true;
    let mut buffer = Vec::with_capacity(256);
    let mut in_sheet_data = false;
    let mut current_row = 0_u32;
    let mut last_row = None;
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"sheetData" => {
                in_sheet_data = true;
            }
            Event::Start(element) if in_sheet_data && element.local_name().as_ref() == b"row" => {
                let row = attributes(&element, reader.decoder())?
                    .get("r")
                    .map_or(Ok(current_row), |value| parse_row_number(value))?;
                current_row = row;
                last_row = Some(row);
            }
            Event::End(element) if element.local_name().as_ref() == b"row" => {
                current_row = current_row.saturating_add(1);
            }
            Event::End(element) if element.local_name().as_ref() == b"sheetData" => {
                return Ok(last_row);
            }
            Event::Eof => {
                return Err(ExcelError::Format(
                    "unexpected end of XML in worksheet".to_owned(),
                ));
            }
            _ => {}
        }
    }
}

fn parse_row_number(value: &str) -> Result<u32> {
    let one_based = value.parse::<u32>().map_err(format_error)?;
    if !(1..=MAX_XLSX_ROW_NUMBER).contains(&one_based) {
        return Err(ExcelError::Format(format!(
            "row index exceeds XLSX limits: {value}"
        )));
    }
    Ok(one_based - 1)
}

fn relationship_part_name(path: &str) -> String {
    path.rsplit_once('/').map_or_else(
        || format!("_rels/{path}.rels"),
        |(directory, file)| format!("{directory}/_rels/{file}.rels"),
    )
}

fn resolve_target(base_part: &str, target: &str) -> Result<String> {
    let candidate = if let Some(absolute) = target.strip_prefix('/') {
        absolute.to_owned()
    } else if let Some((directory, _)) = base_part.rsplit_once('/') {
        format!("{directory}/{target}")
    } else {
        target.to_owned()
    };
    normalize_path(&candidate)
}

fn normalize_path(path: &str) -> Result<String> {
    let mut components = Vec::new();
    for component in path.split('/') {
        match component {
            "" | "." => {}
            ".." => {
                if components.pop().is_none() {
                    return Err(ExcelError::Format(format!(
                        "OOXML relationship escapes package root: {path}"
                    )));
                }
            }
            value => components.push(value),
        }
    }
    if components.is_empty() {
        return Err(ExcelError::Format(
            "empty OOXML relationship target".to_owned(),
        ));
    }
    Ok(components.join("/"))
}

fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[cfg(test)]
mod tests;
