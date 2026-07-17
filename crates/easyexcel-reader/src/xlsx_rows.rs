use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, Read, Seek};

use bigdecimal::BigDecimal;
use easyexcel_core::{CellExtra, CellExtraType, ExcelError, Result};
use quick_xml::escape::resolve_predefined_entity;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Decoder, Reader as XmlReader, XmlVersion};
use ssfmt::{DateSystem, FormatOptions, Locale, format, format_code_from_id};
use zip::ZipArchive;

const MAX_XLSX_ROW_NUMBER: u32 = 1_048_576;
const MAX_XLSX_COLUMN_NUMBER: usize = 16_384;

type Relationships = HashMap<String, (String, String)>;
type RawRelationships = HashMap<String, (String, String, bool)>;

trait ReadSeek: Read + Seek {}

impl<T: Read + Seek> ReadSeek for T {}

pub(crate) struct XlsxRowMetadata {
    archive: ZipArchive<Box<dyn ReadSeek>>,
    path_cache: HashMap<String, String>,
    sheet_paths: HashMap<String, String>,
    cell_formats: Vec<XlsxNumberFormat>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum XlsxNumberFormat {
    Builtin(u32),
    Custom(String),
}

impl XlsxNumberFormat {
    fn display(
        &self,
        value: f64,
        date_1904: bool,
        use_scientific_format: bool,
        locale: &Locale,
    ) -> Option<String> {
        if self.is_general() && is_scientific_magnitude(value) {
            return Some(if use_scientific_format {
                java_scientific_format(value, locale.decimal_separator)
            } else {
                java_plain_extreme_format(value)
            });
        }
        let options = FormatOptions {
            date_system: if date_1904 {
                DateSystem::Date1904
            } else {
                DateSystem::Date1900
            },
            locale: locale.clone(),
        };
        match self {
            Self::Builtin(id) => {
                format_code_from_id(*id).and_then(|code| format(value, code, &options).ok())
            }
            Self::Custom(code) => format(value, code, &options).ok(),
        }
        .map(|value| value.trim().to_owned())
    }

    fn is_general(&self) -> bool {
        match self {
            Self::Builtin(id) => *id == 0,
            Self::Custom(code) => code.trim().eq_ignore_ascii_case("general"),
        }
    }
}

pub(crate) struct XlsxDisplayCell {
    pub(crate) position: (u32, usize),
    pub(crate) display_value: Option<String>,
    pub(crate) decimal_value: Option<BigDecimal>,
}

pub(crate) struct XlsxDisplayCellReader<'a> {
    reader: XmlReader<Box<dyn BufRead + 'a>>,
    cell_formats: &'a [XlsxNumberFormat],
    date_1904: bool,
    use_scientific_format: bool,
    locale: Locale,
    row_index: u32,
    column_index: usize,
    buffer: Vec<u8>,
    cell_buffer: Vec<u8>,
}

impl XlsxRowMetadata {
    pub(crate) fn new(input: impl Read + Seek + 'static) -> Result<Self> {
        Self::new_boxed(Box::new(input))
    }

    fn new_boxed(input: Box<dyn ReadSeek>) -> Result<Self> {
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
        let (sheet_paths, _) = read_workbook_metadata(
            &mut archive,
            &path_cache,
            &workbook_path,
            &workbook_relationships,
        )?;
        let cell_formats = workbook_relationships
            .values()
            .find(|(_, relationship_type)| relationship_type.ends_with("/styles"))
            .map(|(target, _)| resolve_target(&workbook_path, target))
            .transpose()?
            .map(|styles_path| read_cell_formats(&mut archive, &path_cache, &styles_path))
            .transpose()?
            .unwrap_or_else(|| vec![XlsxNumberFormat::Builtin(0)]);
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
            cell_formats,
        })
    }

    pub(crate) fn display_cells(
        &mut self,
        sheet_name: &str,
        use_1904_windowing: bool,
        use_scientific_format: bool,
        locale: Locale,
    ) -> Result<XlsxDisplayCellReader<'_>> {
        let path = self
            .sheet_paths
            .get(sheet_name)
            .cloned()
            .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_owned()))?;
        let actual_path = cached_path(&self.path_cache, &path);
        let file = self.archive.by_name(actual_path).map_err(format_error)?;
        let reader = boxed_xml_reader(BufReader::new(file));
        XlsxDisplayCellReader::new(
            reader,
            &self.cell_formats,
            use_1904_windowing,
            use_scientific_format,
            locale,
        )
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

    pub(crate) fn extras(
        &mut self,
        sheet_name: &str,
        enabled: &HashSet<CellExtraType>,
    ) -> Result<Vec<CellExtra>> {
        let sheet_path = self
            .sheet_paths
            .get(sheet_name)
            .cloned()
            .ok_or_else(|| ExcelError::SheetNotFound(sheet_name.to_owned()))?;
        let relationships_path = relationship_part_name(&sheet_path);
        let relationships = if self
            .path_cache
            .contains_key(&relationships_path.to_ascii_lowercase())
        {
            read_raw_relationships(&mut self.archive, &self.path_cache, &relationships_path)?
        } else {
            RawRelationships::new()
        };
        let mut extras = read_worksheet_extras(
            &mut self.archive,
            &self.path_cache,
            &sheet_path,
            &relationships,
            enabled,
        )?;
        if enabled.contains(&CellExtraType::Comment)
            && let Some((target, _, false)) = relationships
                .values()
                .find(|(_, relationship_type, _)| relationship_type.ends_with("/comments"))
        {
            let comments_path = resolve_target(&sheet_path, target)?;
            extras.extend(read_comments(
                &mut self.archive,
                &self.path_cache,
                &comments_path,
            )?);
        }
        Ok(extras)
    }
}

impl<'a> XlsxDisplayCellReader<'a> {
    fn new(
        mut reader: XmlReader<Box<dyn BufRead + 'a>>,
        cell_formats: &'a [XlsxNumberFormat],
        date_1904: bool,
        use_scientific_format: bool,
        locale: Locale,
    ) -> Result<Self> {
        let mut buffer = Vec::with_capacity(256);
        loop {
            buffer.clear();
            match reader.read_event_into(&mut buffer).map_err(format_error)? {
                Event::Start(element) if element.local_name().as_ref() == b"sheetData" => break,
                Event::Eof => {
                    return Err(ExcelError::Format(
                        "unexpected end of XML before worksheet data".to_owned(),
                    ));
                }
                _ => {}
            }
        }
        Ok(Self {
            reader,
            cell_formats,
            date_1904,
            use_scientific_format,
            locale,
            row_index: 0,
            column_index: 0,
            buffer,
            cell_buffer: Vec::with_capacity(256),
        })
    }

    pub(crate) fn next_cell(&mut self) -> Result<Option<XlsxDisplayCell>> {
        loop {
            self.buffer.clear();
            match self
                .reader
                .read_event_into(&mut self.buffer)
                .map_err(format_error)?
            {
                Event::Start(element) if element.local_name().as_ref() == b"row" => {
                    let values = attributes(&element, self.reader.decoder())?;
                    self.row_index = values
                        .get("r")
                        .map_or(Ok(self.row_index), |value| parse_row_number(value))?;
                    self.column_index = 0;
                }
                Event::Start(element) if element.local_name().as_ref() == b"c" => {
                    let values = attributes(&element, self.reader.decoder())?;
                    let position = values
                        .get("r")
                        .map_or(Ok((self.row_index, self.column_index)), |reference| {
                            parse_cell_reference(reference)
                        })?;
                    let style_index = values
                        .get("s")
                        .map(|value| value.parse::<usize>().map_err(format_error))
                        .transpose()?
                        .unwrap_or_default();
                    let numeric = values.get("t").is_none_or(|value| value == "n");
                    let (display_value, decimal_value) =
                        self.read_cell_display(style_index, numeric)?;
                    self.row_index = position.0;
                    self.column_index = position.1.saturating_add(1);
                    return Ok(Some(XlsxDisplayCell {
                        position,
                        display_value,
                        decimal_value,
                    }));
                }
                Event::End(element) if element.local_name().as_ref() == b"row" => {
                    self.row_index = self.row_index.saturating_add(1);
                    self.column_index = 0;
                }
                Event::End(element) if element.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Event::Eof => {
                    return Err(ExcelError::Format(
                        "unexpected end of XML in worksheet data".to_owned(),
                    ));
                }
                _ => {}
            }
        }
    }

    fn read_cell_display(
        &mut self,
        style_index: usize,
        numeric: bool,
    ) -> Result<(Option<String>, Option<BigDecimal>)> {
        let mut in_value = false;
        let mut raw_value = String::new();
        loop {
            self.cell_buffer.clear();
            match self
                .reader
                .read_event_into(&mut self.cell_buffer)
                .map_err(format_error)?
            {
                Event::Start(element) if element.local_name().as_ref() == b"v" => {
                    in_value = true;
                }
                Event::Text(value) if in_value => {
                    raw_value.push_str(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::CData(value) if in_value => {
                    raw_value.push_str(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::End(element) if element.local_name().as_ref() == b"v" => {
                    in_value = false;
                }
                Event::End(element) if element.local_name().as_ref() == b"c" => {
                    if !numeric || raw_value.is_empty() {
                        return Ok((None, None));
                    }
                    let number =
                        excel_display_number(raw_value.parse::<f64>().map_err(format_error)?);
                    if !number.is_finite() {
                        return Err(ExcelError::Format(
                            "non-finite XLSX numeric cell value".to_owned(),
                        ));
                    }
                    let decimal = number
                        .to_string()
                        .parse::<BigDecimal>()
                        .expect("a finite f64 string is always a valid decimal");
                    let display_value = self.cell_formats.get(style_index).and_then(|format| {
                        format.display(
                            number,
                            self.date_1904,
                            self.use_scientific_format,
                            &self.locale,
                        )
                    });
                    return Ok((display_value, Some(decimal)));
                }
                Event::Eof => {
                    return Err(ExcelError::Format(
                        "unexpected end of XML in worksheet cell".to_owned(),
                    ));
                }
                _ => {}
            }
        }
    }
}

fn excel_display_number(value: f64) -> f64 {
    if value == 0.0 || !value.is_finite() {
        return value;
    }
    format!("{value:.14e}").parse().unwrap_or(value)
}

fn is_scientific_magnitude(value: f64) -> bool {
    let absolute = value.abs();
    absolute >= 1E11 || (absolute <= 1E-10 && absolute > 0.0)
}

fn java_plain_extreme_format(value: f64) -> String {
    let rounded = value.round();
    if rounded == 0.0 {
        "0".to_owned()
    } else {
        format!("{rounded:.0}")
    }
}

fn java_scientific_format(value: f64, decimal_separator: char) -> String {
    let formatted = format!("{value:.5e}");
    let (mantissa, exponent) = formatted
        .split_once('e')
        .expect("Rust scientific formatting always contains an exponent");
    let mantissa = mantissa.trim_end_matches('0').trim_end_matches('.');
    let exponent = exponent
        .parse::<i32>()
        .expect("Rust scientific formatting always emits a numeric exponent");
    let mantissa = if decimal_separator == '.' {
        mantissa.to_owned()
    } else {
        mantissa.replace('.', &decimal_separator.to_string())
    };
    format!("{mantissa}E{exponent}")
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

fn boxed_xml_reader<'a>(input: impl BufRead + 'a) -> XmlReader<Box<dyn BufRead + 'a>> {
    let mut reader = XmlReader::from_reader(Box::new(input) as Box<dyn BufRead + 'a>);
    let config = reader.config_mut();
    config.check_end_names = false;
    config.check_comments = false;
    config.expand_empty_elements = true;
    reader
}

fn read_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    path: &str,
) -> Result<Relationships> {
    Ok(read_raw_relationships(archive, cache, path)?
        .into_iter()
        .filter_map(|(id, (target, relationship_type, external))| {
            (!external).then_some((id, (target, relationship_type)))
        })
        .collect())
}

fn read_raw_relationships<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    path: &str,
) -> Result<RawRelationships> {
    let mut reader = xml_reader(archive, cache, path)?;
    let mut relationships = HashMap::new();
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"Relationship" => {
                let attributes = attributes(&element, reader.decoder())?;
                let Some(id) = attributes.get("Id") else {
                    continue;
                };
                let target = attributes.get("Target").cloned().unwrap_or_default();
                let relationship_type = attributes.get("Type").cloned().unwrap_or_default();
                let external = attributes
                    .get("TargetMode")
                    .is_some_and(|mode| mode == "External");
                relationships.insert(id.clone(), (target, relationship_type, external));
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

fn read_worksheet_extras<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    sheet_path: &str,
    relationships: &RawRelationships,
    enabled: &HashSet<CellExtraType>,
) -> Result<Vec<CellExtra>> {
    let file = archive
        .by_name(cached_path(cache, sheet_path))
        .map_err(format_error)?;
    let mut input = BufReader::new(file);
    parse_worksheet_extras(&mut input, sheet_path, relationships, enabled)
}

fn parse_worksheet_extras(
    input: &mut dyn BufRead,
    sheet_path: &str,
    relationships: &RawRelationships,
    enabled: &HashSet<CellExtraType>,
) -> Result<Vec<CellExtra>> {
    let mut reader = XmlReader::from_reader(input);
    let config = reader.config_mut();
    config.check_end_names = false;
    config.check_comments = false;
    config.expand_empty_elements = true;
    let mut extras = Vec::new();
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element)
                if enabled.contains(&CellExtraType::Merge)
                    && element.local_name().as_ref() == b"mergeCell" =>
            {
                let attributes = attributes(&element, reader.decoder())?;
                let reference = required_attribute(&attributes, "ref", "merge cell")?;
                let (first_row, last_row, first_column, last_column) = parse_cell_range(reference)?;
                extras.push(CellExtra::new(
                    CellExtraType::Merge,
                    None,
                    first_row,
                    last_row,
                    first_column,
                    last_column,
                ));
            }
            Event::Start(element)
                if enabled.contains(&CellExtraType::Hyperlink)
                    && element.local_name().as_ref() == b"hyperlink" =>
            {
                let attributes = attributes(&element, reader.decoder())?;
                let reference = required_attribute(&attributes, "ref", "hyperlink")?;
                let text = hyperlink_text(&attributes, relationships)?;
                let (first_row, last_row, first_column, last_column) = parse_cell_range(reference)?;
                extras.push(CellExtra::new(
                    CellExtraType::Hyperlink,
                    Some(text),
                    first_row,
                    last_row,
                    first_column,
                    last_column,
                ));
            }
            Event::End(element) if element.local_name().as_ref() == b"worksheet" => break,
            Event::Eof => {
                return Err(ExcelError::Format(format!(
                    "unexpected end of XML in {sheet_path}"
                )));
            }
            _ => {}
        }
    }
    Ok(extras)
}

fn hyperlink_text(
    attributes: &HashMap<String, String>,
    relationships: &RawRelationships,
) -> Result<String> {
    if let Some(location) = attributes.get("location") {
        return Ok(location.clone());
    }
    let relationship_id = required_attribute(attributes, "id", "hyperlink")?;
    relationships
        .get(relationship_id)
        .filter(|(_, relationship_type, _)| relationship_type.ends_with("/hyperlink"))
        .map(|(target, _, _)| target.clone())
        .ok_or_else(|| {
            ExcelError::Format(format!(
                "hyperlink relationship not found: {relationship_id}"
            ))
        })
}

fn read_comments<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    comments_path: &str,
) -> Result<Vec<CellExtra>> {
    let file = archive
        .by_name(cached_path(cache, comments_path))
        .map_err(format_error)?;
    let mut input = BufReader::new(file);
    parse_comments(&mut input, comments_path)
}

fn parse_comments(input: &mut dyn BufRead, comments_path: &str) -> Result<Vec<CellExtra>> {
    let mut reader = XmlReader::from_reader(input);
    let config = reader.config_mut();
    config.check_end_names = false;
    config.check_comments = false;
    config.expand_empty_elements = true;
    let mut extras = Vec::new();
    let mut buffer = Vec::with_capacity(256);
    let mut current = None;
    let mut text = String::new();
    let mut in_text_run = false;
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"comment" => {
                let attributes = attributes(&element, reader.decoder())?;
                let reference = required_attribute(&attributes, "ref", "comment")?;
                current = Some(parse_cell_range(reference)?);
                text.clear();
            }
            Event::Start(element) if current.is_some() && element.local_name().as_ref() == b"t" => {
                in_text_run = true;
            }
            Event::Text(value) if in_text_run => {
                let decoded = value
                    .xml_content(XmlVersion::Implicit1_0)
                    .map_err(format_error)?;
                text.push_str(&decoded);
            }
            Event::CData(value) if in_text_run => {
                text.push_str(
                    &value
                        .xml_content(XmlVersion::Implicit1_0)
                        .map_err(format_error)?,
                );
            }
            Event::GeneralRef(value) if in_text_run => {
                if let Some(character) = value.resolve_char_ref().map_err(format_error)? {
                    text.push(character);
                } else {
                    // XML entity names are ASCII by specification. Lossy decoding keeps
                    // malformed input on the typed unrecognized-entity path as well.
                    let name = String::from_utf8_lossy(value.as_ref());
                    let replacement = resolve_predefined_entity(&name).ok_or_else(|| {
                        ExcelError::Format(format!("unrecognized XML entity: {name}"))
                    })?;
                    text.push_str(replacement);
                }
            }
            Event::End(element) if element.local_name().as_ref() == b"t" => {
                in_text_run = false;
            }
            Event::End(element) if element.local_name().as_ref() == b"comment" => {
                let (first_row, last_row, first_column, last_column) = current
                    .take()
                    .ok_or_else(|| ExcelError::Format("comment start is missing".to_owned()))?;
                extras.push(CellExtra::new(
                    CellExtraType::Comment,
                    Some(text.clone()),
                    first_row,
                    last_row,
                    first_column,
                    last_column,
                ));
            }
            Event::End(element) if element.local_name().as_ref() == b"comments" => break,
            Event::Eof => {
                return Err(ExcelError::Format(format!(
                    "unexpected end of XML in {comments_path}"
                )));
            }
            _ => {}
        }
    }
    Ok(extras)
}

fn required_attribute<'a>(
    attributes: &'a HashMap<String, String>,
    name: &str,
    element: &str,
) -> Result<&'a str> {
    attributes
        .get(name)
        .map(String::as_str)
        .ok_or_else(|| ExcelError::Format(format!("{element} {name} is missing")))
}

fn parse_cell_range(reference: &str) -> Result<(u32, u32, usize, usize)> {
    let (first, last) = reference
        .split_once(':')
        .map_or((reference, reference), |range| range);
    let (first_row, first_column) = parse_cell_reference(first)?;
    let (last_row, last_column) = parse_cell_reference(last)?;
    if first_row > last_row || first_column > last_column {
        return Err(ExcelError::Format(format!(
            "invalid cell range ordering: {reference}"
        )));
    }
    Ok((first_row, last_row, first_column, last_column))
}

fn parse_cell_reference(reference: &str) -> Result<(u32, usize)> {
    let reference = reference.strip_prefix('$').unwrap_or(reference);
    let column_end = reference
        .find(|character: char| !character.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (column, row) = reference.split_at(column_end);
    let row = row.strip_prefix('$').unwrap_or(row);
    if column.is_empty() || row.is_empty() || !row.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(ExcelError::Format(format!(
            "invalid cell reference: {reference}"
        )));
    }
    let mut one_based_column = 0_usize;
    for letter in column.bytes() {
        one_based_column = one_based_column
            .checked_mul(26)
            .and_then(|value| {
                value.checked_add(usize::from(letter.to_ascii_uppercase() - b'A' + 1))
            })
            .ok_or_else(|| ExcelError::Format(format!("invalid cell reference: {reference}")))?;
    }
    if !(1..=MAX_XLSX_COLUMN_NUMBER).contains(&one_based_column) {
        return Err(ExcelError::Format(format!(
            "column index exceeds XLSX limits: {reference}"
        )));
    }
    Ok((parse_row_number(row)?, one_based_column - 1))
}

fn read_cell_formats<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    styles_path: &str,
) -> Result<Vec<XlsxNumberFormat>> {
    let mut reader = xml_reader(archive, cache, styles_path)?;
    let mut custom_formats = HashMap::new();
    let mut cell_formats = Vec::new();
    let mut in_cell_formats = false;
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"numFmt" => {
                let values = attributes(&element, reader.decoder())?;
                if let (Some(id), Some(code)) = (values.get("numFmtId"), values.get("formatCode")) {
                    custom_formats.insert(id.parse::<u32>().map_err(format_error)?, code.clone());
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"cellXfs" => {
                in_cell_formats = true;
            }
            Event::Start(element) if in_cell_formats && element.local_name().as_ref() == b"xf" => {
                let values = attributes(&element, reader.decoder())?;
                let id = values
                    .get("numFmtId")
                    .map(|value| value.parse::<u32>().map_err(format_error))
                    .transpose()?
                    .unwrap_or_default();
                cell_formats.push(custom_formats.get(&id).map_or_else(
                    || XlsxNumberFormat::Builtin(id),
                    |code| XlsxNumberFormat::Custom(code.clone()),
                ));
            }
            Event::End(element) if element.local_name().as_ref() == b"cellXfs" => {
                in_cell_formats = false;
            }
            Event::End(element) if element.local_name().as_ref() == b"styleSheet" => break,
            Event::Eof => {
                return Err(ExcelError::Format(
                    "unexpected end of XML in styles".to_owned(),
                ));
            }
            _ => {}
        }
    }
    if cell_formats.is_empty() {
        cell_formats.push(XlsxNumberFormat::Builtin(0));
    }
    Ok(cell_formats)
}

fn read_workbook_metadata<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    workbook_path: &str,
    relationships: &Relationships,
) -> Result<(HashMap<String, String>, bool)> {
    let mut reader = xml_reader(archive, cache, workbook_path)?;
    let mut sheets = HashMap::new();
    let mut date_1904 = false;
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"workbookPr" => {
                let values = attributes(&element, reader.decoder())?;
                date_1904 = values
                    .get("date1904")
                    .is_some_and(|value| matches!(value.as_str(), "1" | "true"));
            }
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
    Ok((sheets, date_1904))
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
