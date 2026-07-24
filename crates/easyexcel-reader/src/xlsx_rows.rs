use std::collections::{HashMap, HashSet};
use std::io::{BufRead, BufReader, Read, Seek};

use bigdecimal::BigDecimal;
use calamine::{ExcelDateTime, ExcelDateTimeType};
use easyexcel_core::constant::builtin_format_code;
use easyexcel_core::metadata::format::{
    java_compat_date_format_code, java_compat_display, java_compat_format_code,
};
use easyexcel_core::{CellValue, ExcelError, FormulaData, Result};
use quick_xml::escape::resolve_predefined_entity;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Decoder, Reader as XmlReader, XmlVersion};
use ssfmt::{DateSystem, FormatOptions, Locale, NumberFormat, format, format_code_from_id};
use zip::ZipArchive;

use crate::ReadOptions;
use crate::analysis::v07::handlers::cell_formula_tag_handler::CellFormulaTagHandler;
use crate::analysis::v07::handlers::cell_inline_string_value_tag_handler::CellInlineStringValueTagHandler;
use crate::analysis::v07::handlers::cell_tag_handler::CellTagHandler;
use crate::analysis::v07::handlers::cell_value_tag_handler::CellValueTagHandler;
use crate::analysis::v07::handlers::count_tag_handler::CountTagHandler;
use crate::analysis::v07::handlers::hyperlink_tag_handler::HyperlinkTagHandler;
use crate::analysis::v07::handlers::merge_cell_tag_handler::MergeCellTagHandler;
use crate::analysis::v07::handlers::row_tag_handler::RowTagHandler;
use crate::analysis::v07::handlers::sax::shared_strings_table_handler::{
    SharedStringsTableHandler, local_tag, utf_decode,
};
use crate::analysis::v07::handlers::xlsx_tag_handler::XlsxTagHandler;
use crate::cache::resolve_read_cache_mode;
use crate::read_cache::{
    ReadCacheMode, SharedStringCache, SharedStringCacheReader, SharedStringCacheWriter,
    create_cache, memory_cache,
};

/// Prefer EasyExcel BuiltinFormats over ssfmt ECMA table (Java locale-aware codes).
fn easyexcel_builtin_format_code(id: u32) -> Option<&'static str> {
    builtin_format_code(id as u16)
}

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
    sheet_names: Vec<String>,
    cell_formats: Vec<XlsxNumberFormat>,
    shared_strings: Box<dyn SharedStringCacheReader>,
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
                // Prefer EasyExcel BuiltinFormats (locale-aware CN/ALL tables) over ssfmt's
                // ECMA builtin table so STRING display matches Java (e.g. id 22 → yyyy-m-d h:mm).
                let code = easyexcel_builtin_format_code(*id).or_else(|| format_code_from_id(*id));
                code.and_then(|code| format_with_resolved_code(value, code, &options))
            }
            Self::Custom(code) => format_with_resolved_code(value, code, &options),
        }
    }

    fn is_general(&self) -> bool {
        match self {
            Self::Builtin(id) => *id == 0,
            Self::Custom(code) => code.trim().eq_ignore_ascii_case("general"),
        }
    }

    fn is_date_format(&self) -> bool {
        let code = match self {
            Self::Builtin(id) => {
                easyexcel_builtin_format_code(*id).or_else(|| format_code_from_id(*id))
            }
            Self::Custom(code) => Some(code.as_str()),
        };
        code.and_then(|code| NumberFormat::parse(code).ok())
            .is_some_and(|format| format.is_date_format())
    }
}

pub(crate) struct XlsxDisplayCell {
    pub(crate) position: (u32, usize),
    pub(crate) value: CellValue,
    pub(crate) formula: Option<FormulaData>,
    pub(crate) display_value: Option<String>,
    pub(crate) decimal_value: Option<BigDecimal>,
}

type ParsedCell = (
    CellValue,
    Option<FormulaData>,
    Option<String>,
    Option<BigDecimal>,
);

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
    shared_strings: &'a dyn SharedStringCacheReader,
}

impl XlsxRowMetadata {
    #[cfg(test)]
    pub(crate) fn new(input: impl Read + Seek + 'static) -> Result<Self> {
        Self::new_with_cache(input, &ReadOptions::default())
    }

    pub(crate) fn new_with_cache(
        input: impl Read + Seek + 'static,
        options: &ReadOptions,
    ) -> Result<Self> {
        Self::new_boxed(Box::new(input), options)
    }

    fn new_boxed(input: Box<dyn ReadSeek>, options: &ReadOptions) -> Result<Self> {
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
        let (sheets, _) = read_workbook_metadata(
            &mut archive,
            &path_cache,
            &workbook_path,
            &workbook_relationships,
        )?;
        let sheet_names = sheets.iter().map(|(name, _)| name.clone()).collect();
        let sheet_paths = sheets.into_iter().collect::<HashMap<_, _>>();
        let cell_formats = workbook_relationships
            .values()
            .find(|(_, relationship_type)| relationship_type.ends_with("/styles"))
            .map(|(target, _)| resolve_target(&workbook_path, target))
            .transpose()?
            .map(|styles_path| read_cell_formats(&mut archive, &path_cache, &styles_path))
            .transpose()?
            .unwrap_or_else(|| vec![XlsxNumberFormat::Builtin(0)]);
        let shared_strings_path = workbook_relationships
            .values()
            .find(|(_, relationship_type)| relationship_type.ends_with("/sharedStrings"))
            .map(|(target, _)| resolve_target(&workbook_path, target))
            .transpose()?;
        let shared_strings = match shared_strings_path {
            Some(path) => read_shared_strings(&mut archive, &path_cache, &path, options)?,
            None => memory_cache(),
        };
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
            sheet_names,
            cell_formats,
            shared_strings,
        })
    }

    pub(crate) fn sheet_names(&self) -> Vec<String> {
        self.sheet_names.clone()
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
            self.shared_strings.as_ref(),
        )
    }

    pub(crate) fn last_explicit_row(&mut self, sheet_name: &str) -> Result<Option<u32>> {
        // Uses scan_last_row → CountTagHandler for `<dimension>` fallback
        // (Java `com.alibaba.excel.analysis.v07.handlers.CountTagHandler`).
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
        enabled: &HashSet<easyexcel_core::CellExtraType>,
    ) -> Result<Vec<easyexcel_core::CellExtra>> {
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
        if enabled.contains(&easyexcel_core::CellExtraType::Comment)
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
        shared_strings: &'a dyn SharedStringCacheReader,
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
            shared_strings,
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
                    // Java `RowTagHandler.startElement` → `PositionUtils.getRowByRowTagt`
                    self.row_index = RowTagHandler::resolve_row_index(
                        values.get("r").map(String::as_str),
                        self.row_index,
                    )?;
                    self.column_index = 0;
                }
                Event::Start(element) if element.local_name().as_ref() == b"c" => {
                    let values = attributes(&element, self.reader.decoder())?;
                    // Java `CellTagHandler.startElement` attribute portion
                    let started =
                        CellTagHandler::parse_start(&values, self.row_index, self.column_index)?;
                    let (value, formula, display_value, decimal_value) =
                        self.read_cell(started.style_index, started.cell_type.as_deref())?;
                    self.row_index = started.position.0;
                    self.column_index = started.position.1.saturating_add(1);
                    return Ok(Some(XlsxDisplayCell {
                        position: started.position,
                        value,
                        formula,
                        display_value,
                        decimal_value,
                    }));
                }
                Event::End(element) if element.local_name().as_ref() == b"row" => {
                    // Java `RowTagHandler.endElement` advances the cursor after
                    // emitting the row; display-cell streaming only needs the index.
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

    fn read_cell(&mut self, style_index: usize, cell_type: Option<&str>) -> Result<ParsedCell> {
        // Dual-track: accumulate via Java-parity handlers while the quick_xml
        // loop remains the event driver (只增不减).
        let mut value_handler = CellValueTagHandler::new();
        let mut inline_handler = CellInlineStringValueTagHandler::new();
        let mut formula_handler = CellFormulaTagHandler::new();
        let mut in_value = false;
        let mut in_formula = false;
        let mut in_text = false;
        let mut phonetic_depth = 0_u32;
        loop {
            self.cell_buffer.clear();
            match self
                .reader
                .read_event_into(&mut self.cell_buffer)
                .map_err(format_error)?
            {
                Event::Start(element) if element.local_name().as_ref() == b"v" => {
                    // Java `CellValueTagHandler` / `AbstractCellValueTagHandler`
                    in_value = true;
                }
                Event::Start(element) if element.local_name().as_ref() == b"f" => {
                    // Java `CellFormulaTagHandler.startElement`
                    formula_handler.begin_formula();
                    in_formula = true;
                }
                Event::Start(element) if element.local_name().as_ref() == b"rPh" => {
                    phonetic_depth = phonetic_depth.saturating_add(1);
                }
                Event::Start(element)
                    if phonetic_depth == 0 && element.local_name().as_ref() == b"t" =>
                {
                    // Java `CellInlineStringValueTagHandler` (inherits characters)
                    in_text = true;
                }
                Event::Text(value) if in_value => {
                    value_handler.characters(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::Text(value) if in_formula => {
                    // Java `CellFormulaTagHandler.characters`
                    formula_handler.characters(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::Text(value) if in_text => {
                    inline_handler.characters(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::CData(value) if in_value => {
                    value_handler.characters(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::CData(value) if in_formula => {
                    formula_handler.characters(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::CData(value) if in_text => {
                    inline_handler.characters(
                        &value
                            .xml_content(XmlVersion::Implicit1_0)
                            .map_err(format_error)?,
                    );
                }
                Event::End(element) if element.local_name().as_ref() == b"v" => {
                    in_value = false;
                }
                Event::End(element) if element.local_name().as_ref() == b"f" => {
                    // Java `CellFormulaTagHandler.endElement` attaches formula to
                    // tempCellData; we keep it on the handler until `</c>`.
                    in_formula = false;
                }
                Event::End(element) if element.local_name().as_ref() == b"t" => {
                    in_text = false;
                }
                Event::End(element) if element.local_name().as_ref() == b"rPh" => {
                    phonetic_depth = phonetic_depth.saturating_sub(1);
                }
                Event::End(element) if element.local_name().as_ref() == b"c" => {
                    let formula_text = formula_handler.finish_formula();
                    let formula =
                        (!formula_text.is_empty()).then(|| FormulaData::new(formula_text));
                    let raw_value = value_handler.take();
                    let inline_value = inline_handler.take();
                    return self.finish_cell(
                        style_index,
                        cell_type,
                        &raw_value,
                        &inline_value,
                        formula,
                    );
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

    fn finish_cell(
        &mut self,
        style_index: usize,
        cell_type: Option<&str>,
        raw_value: &str,
        inline_value: &str,
        formula: Option<FormulaData>,
    ) -> Result<ParsedCell> {
        let number = if matches!(cell_type, Some("n") | None) && !raw_value.is_empty() {
            let number = excel_display_number(raw_value.parse::<f64>().map_err(format_error)?);
            if !number.is_finite() {
                return Err(ExcelError::Format(
                    "non-finite XLSX numeric cell value".to_owned(),
                ));
            }
            Some(number)
        } else {
            None
        };
        let value = match cell_type {
            Some("s") => {
                if raw_value.is_empty() {
                    return Ok((CellValue::Empty, formula, None, None));
                }
                let index = raw_value.parse::<usize>().map_err(format_error)?;
                CellValue::String(self.shared_strings.get(index)?)
            }
            Some("inlineStr" | "str") => {
                CellValue::String(utf_decode(if inline_value.is_empty() {
                    raw_value
                } else {
                    inline_value
                }))
            }
            Some("b") => CellValue::Bool(matches!(raw_value, "1" | "true")),
            Some("e") => CellValue::Error(raw_value.to_owned()),
            Some("d") => CellValue::String(raw_value.to_owned()),
            Some("n") | None => {
                if raw_value.is_empty() {
                    CellValue::Empty
                } else {
                    self.numeric_cell(style_index, number.expect("numeric value was parsed"))
                }
            }
            Some(other) => {
                return Err(ExcelError::Format(format!(
                    "unsupported XLSX cell type: {other}"
                )));
            }
        };
        let (display_value, decimal_value) = if let Some(number) = number {
            let decimal = number
                .to_string()
                .parse::<BigDecimal>()
                .expect("a finite f64 string is always a valid decimal");
            let display = self.cell_formats.get(style_index).and_then(|format| {
                format.display(
                    number,
                    self.date_1904,
                    self.use_scientific_format,
                    &self.locale,
                )
            });
            (display, Some(decimal))
        } else {
            (None, None)
        };
        Ok((value, formula, display_value, decimal_value))
    }

    fn numeric_cell(&self, style_index: usize, number: f64) -> CellValue {
        if self
            .cell_formats
            .get(style_index)
            .is_some_and(XlsxNumberFormat::is_date_format)
        {
            let date = ExcelDateTime::new(number, ExcelDateTimeType::DateTime, self.date_1904);
            return super::excel_datetime_cell(&date, self.date_1904);
        }
        CellValue::Float(number)
    }
}

/// Format a numeric cell with an Excel format code (BuiltinFormats / custom).
pub(crate) fn format_with_code(
    value: f64,
    code: &str,
    date_1904: bool,
    locale: &Locale,
) -> Option<String> {
    let options = FormatOptions {
        date_system: if date_1904 {
            DateSystem::Date1904
        } else {
            DateSystem::Date1900
        },
        locale: locale.clone(),
    };
    format_with_resolved_code(value, code, &options)
}

/// Apply EasyExcel number-format cleaning then ssfmt + [`java_compat_display`].
///
/// Date codes go through [`java_compat_date_format_code`] (CN `上午/下午` → `AM/PM`,
/// `mmmmm` → POI PUA wrap) while keeping escaped literals (`yyyy\-m\-dd`).
/// Number codes go through [`java_compat_format_code`] so `_X` pads disappear and
/// `\ ` trailing spaces remain.
fn format_with_resolved_code(value: f64, code: &str, options: &FormatOptions) -> Option<String> {
    let is_date = NumberFormat::parse(code)
        .ok()
        .is_some_and(|parsed| parsed.is_date_format());
    let resolved = if is_date {
        java_compat_date_format_code(code)
    } else {
        java_compat_format_code(code)
    };
    format(value, &resolved, options)
        .ok()
        .map(|formatted| java_compat_display(&formatted))
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
    enabled: &HashSet<easyexcel_core::CellExtraType>,
) -> Result<Vec<easyexcel_core::CellExtra>> {
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
    enabled: &HashSet<easyexcel_core::CellExtraType>,
) -> Result<Vec<easyexcel_core::CellExtra>> {
    let mut reader = XmlReader::from_reader(input);
    let config = reader.config_mut();
    config.check_end_names = false;
    config.check_comments = false;
    config.expand_empty_elements = true;
    let mut extras = Vec::new();
    let mut merge_handler =
        MergeCellTagHandler::new(enabled.contains(&easyexcel_core::CellExtraType::Merge));
    let mut hyperlink_handler =
        HyperlinkTagHandler::new(enabled.contains(&easyexcel_core::CellExtraType::Hyperlink));
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"mergeCell" => {
                // Java `MergeCellTagHandler.startElement`
                let attributes = attributes(&element, reader.decoder())?;
                merge_handler.start_merge_required(&attributes)?;
                if let Some(extra) = merge_handler.last_extra.take() {
                    extras.push(extra);
                }
            }
            Event::Start(element) if element.local_name().as_ref() == b"hyperlink" => {
                // Java `HyperlinkTagHandler.startElement`
                let attributes = attributes(&element, reader.decoder())?;
                hyperlink_handler.start_hyperlink_required(&attributes, &|r_id| {
                    relationships
                        .get(r_id)
                        .filter(|(_, relationship_type, _)| {
                            relationship_type.ends_with("/hyperlink")
                        })
                        .map(|(target, _, _)| target.clone())
                        .ok_or_else(|| {
                            ExcelError::Format(format!("hyperlink relationship not found: {r_id}"))
                        })
                })?;
                if let Some(extra) = hyperlink_handler.last_extra.take() {
                    extras.push(extra);
                }
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

fn read_comments<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    comments_path: &str,
) -> Result<Vec<easyexcel_core::CellExtra>> {
    let file = archive
        .by_name(cached_path(cache, comments_path))
        .map_err(format_error)?;
    let mut input = BufReader::new(file);
    parse_comments(&mut input, comments_path)
}

fn parse_comments(
    input: &mut dyn BufRead,
    comments_path: &str,
) -> Result<Vec<easyexcel_core::CellExtra>> {
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
                extras.push(easyexcel_core::CellExtra::new(
                    easyexcel_core::CellExtraType::Comment,
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

fn read_shared_strings<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    path_cache: &HashMap<String, String>,
    path: &str,
    options: &ReadOptions,
) -> Result<Box<dyn SharedStringCacheReader>> {
    let mode = options.read_cache;
    let selector = options
        .read_cache_selector
        .as_ref()
        .map(|stored| stored as &dyn crate::cache::ReadCacheSelector);
    let mut cache_factory = |xml_size| {
        let effective = resolve_read_cache_mode(mode, selector, xml_size);
        create_cache(effective, xml_size)
    };
    read_shared_strings_with_factory(archive, path_cache, path, &mut cache_factory)
}

fn read_shared_strings_with_factory<R>(
    archive: &mut ZipArchive<R>,
    path_cache: &HashMap<String, String>,
    path: &str,
    cache_factory: &mut dyn FnMut(u64) -> Result<Box<dyn SharedStringCache>>,
) -> Result<Box<dyn SharedStringCacheReader>>
where
    R: Read + Seek,
{
    let actual_path = cached_path(path_cache, path);
    let file = archive.by_name(actual_path).map_err(format_error)?;
    let xml_size = file.size();
    let mut cache = cache_factory(xml_size)?;
    let mut input = BufReader::new(file);
    parse_shared_strings(&mut input, cache.as_mut())?;
    // After writing is complete, convert to read-only reader for concurrent access
    cache.finish()
}

fn parse_shared_strings(
    input: &mut dyn BufRead,
    cache: &mut dyn SharedStringCacheWriter,
) -> Result<()> {
    let mut reader = XmlReader::from_reader(input);
    reader.config_mut().expand_empty_elements = true;
    let mut buffer = Vec::with_capacity(256);
    // Java `SharedStringsTableHandler` — driven by the same quick_xml loop.
    let mut handler = SharedStringsTableHandler::new();
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) => {
                let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                handler.start_element(local_tag(&name));
            }
            Event::Text(value) => {
                handler.characters(
                    &value
                        .xml_content(XmlVersion::Implicit1_0)
                        .map_err(format_error)?,
                );
            }
            Event::CData(value) => {
                handler.characters(
                    &value
                        .xml_content(XmlVersion::Implicit1_0)
                        .map_err(format_error)?,
                );
            }
            Event::End(element) => {
                let name = String::from_utf8_lossy(element.local_name().as_ref()).into_owned();
                if let Some(decoded) = handler.end_element(local_tag(&name)) {
                    cache.put(decoded)?;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}

fn read_workbook_metadata<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    cache: &HashMap<String, String>,
    workbook_path: &str,
    relationships: &Relationships,
) -> Result<(Vec<(String, String)>, bool)> {
    let mut reader = xml_reader(archive, cache, workbook_path)?;
    let mut sheets = Vec::new();
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
                    sheets.push((name.clone(), resolve_target(workbook_path, target)?));
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

/// Scans worksheet XML for the last explicit row index.
///
/// Prefers real `<row>` markers inside `<sheetData>`. When the sheet has no
/// row elements, falls back to
/// [`CountTagHandler`](crate::analysis::v07::handlers::count_tag_handler::CountTagHandler)
/// parsing of `<dimension ref="...">` (Java
/// `com.alibaba.excel.analysis.v07.handlers.CountTagHandler` →
/// `ReadSheetHolder.approximateTotalRowNumber`).
fn scan_last_row<R: BufRead>(input: R) -> Result<Option<u32>> {
    let mut reader = XmlReader::from_reader(input);
    reader.config_mut().expand_empty_elements = true;
    let mut buffer = Vec::with_capacity(256);
    let mut in_sheet_data = false;
    let mut current_row = 0_u32;
    let mut last_row = None;
    // Java CountTagHandler — dimension-derived approximate total / last row.
    let mut count_handler = CountTagHandler::new();
    loop {
        buffer.clear();
        match reader.read_event_into(&mut buffer).map_err(format_error)? {
            Event::Start(element) if element.local_name().as_ref() == b"dimension" => {
                // Wire CountTagHandler into the last_explicit_row hot path.
                let attrs = attributes(&element, reader.decoder())?;
                count_handler.start_dimension(&attrs)?;
            }
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
                // Prefer observed rows; otherwise use CountTagHandler dimension.
                if last_row.is_some() {
                    return Ok(last_row);
                }
                return Ok(count_handler.last_explicit_row_index());
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
