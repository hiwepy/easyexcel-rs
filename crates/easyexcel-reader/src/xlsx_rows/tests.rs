use std::io::{Cursor, Write};
use std::sync::Arc;

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use super::*;
use crate::XlsxInput;

fn package(entries: &[(&str, &str)]) -> Cursor<Vec<u8>> {
    let mut output = Cursor::new(Vec::new());
    {
        let mut writer = ZipWriter::new(&mut output);
        for (name, contents) in entries {
            writer
                .start_file(*name, SimpleFileOptions::default())
                .expect("start package entry");
            writer
                .write_all(contents.as_bytes())
                .expect("write package entry");
        }
        writer.finish().expect("finish package");
    }
    output.set_position(0);
    output
}

fn package_bytes(entries: &[(&str, &[u8])]) -> Cursor<Vec<u8>> {
    let mut output = Cursor::new(Vec::new());
    {
        let mut writer = ZipWriter::new(&mut output);
        for (name, contents) in entries {
            writer
                .start_file(*name, SimpleFileOptions::default())
                .expect("start package entry");
            writer.write_all(contents).expect("write package entry");
        }
        writer.finish().expect("finish package");
    }
    output.set_position(0);
    output
}

fn archive(entries: &[(&str, &str)]) -> ZipArchive<Box<dyn ReadSeek>> {
    ZipArchive::new(Box::new(package(entries)) as Box<dyn ReadSeek>).expect("valid ZIP archive")
}

fn xlsx_input_archive(entries: &[(&str, &str)]) -> ZipArchive<XlsxInput> {
    ZipArchive::new(xlsx_input(entries)).expect("valid XLSX input archive")
}

fn xlsx_input(entries: &[(&str, &str)]) -> XlsxInput {
    let bytes: Arc<[u8]> = Arc::from(package(entries).into_inner());
    XlsxInput::Memory(Cursor::new(bytes))
}

fn display_xml_reader<'a, R: Read + Seek>(
    archive: &'a mut ZipArchive<R>,
    path: &str,
) -> Result<XmlReader<Box<dyn BufRead + 'a>>> {
    let file = archive.by_name(path).map_err(format_error)?;
    Ok(boxed_xml_reader(BufReader::new(file)))
}

fn valid_entries() -> Vec<(&'static str, &'static str)> {
    vec![
        (
            "_RELS/.RELS",
            r#"<Relationships>
<Relationship Id="external" Type="urn:test" Target="https://example.com" TargetMode="External"/>
<Relationship Type="urn:ignored" Target="ignored"/>
<Relationship Id="office" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="/custom/Workbook.XML"/>
</Relationships>"#,
        ),
        (
            "custom/_rels/Workbook.XML.rels",
            r#"<Relationships>
<Relationship Id="sheet" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="sheets/First.XML"/>
<Relationship Id="chart" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet" Target="charts/chart1.xml"/>
</Relationships>"#,
        ),
        (
            "custom/Workbook.XML",
            r#"<workbook xmlns:r="urn:r"><sheets>
<sheet name="A&amp;B" r:id="sheet"/>
<sheet name="Chart" r:id="chart"/>
</sheets></workbook>"#,
        ),
        (
            "custom/sheets/First.XML",
            r#"<worksheet><ignored/><sheetData><row r="2"/><row/><row r="5"/></sheetData></worksheet>"#,
        ),
    ]
}

#[test]
fn metadata_resolves_strict_case_insensitive_relationships_and_rows() -> Result<()> {
    let mut metadata = XlsxRowMetadata::new(package(&valid_entries()))?;
    assert!(
        metadata
            .display_cells("Missing", false, false, Locale::default())
            .is_err()
    );
    assert_eq!(metadata.last_explicit_row("A&B")?, Some(4));
    assert!(metadata.last_explicit_row("Chart").is_err());
    assert!(metadata.last_explicit_row("Missing").is_err());

    let mut missing_part = XlsxRowMetadata {
        archive: archive(&[]),
        path_cache: HashMap::new(),
        sheet_names: vec!["MissingPart".to_owned()],
        sheet_paths: HashMap::from([("MissingPart".to_owned(), "missing.xml".to_owned())]),
        cell_formats: vec![XlsxNumberFormat::Builtin(0)],
        shared_strings: create_cache(ReadCacheMode::Memory, 0)?,
    };
    assert!(
        missing_part
            .display_cells("MissingPart", false, false, Locale::default())
            .is_err()
    );
    assert!(missing_part.last_explicit_row("MissingPart").is_err());

    let mut missing_sheet_entries = valid_entries();
    missing_sheet_entries.pop();
    assert!(XlsxRowMetadata::new(package(&missing_sheet_entries)).is_err());
    Ok(())
}

#[test]
fn display_cell_stream_formats_exact_numbers_and_tracks_sparse_coordinates() -> Result<()> {
    let formats = vec![
        XlsxNumberFormat::Builtin(0),
        XlsxNumberFormat::Custom("0.00_ ".to_owned()),
    ];
    assert_eq!(
        formats[0]
            .display(1.0, false, false, &Locale::default())
            .as_deref(),
        Some("1")
    );
    assert_eq!(
        formats[1]
            .display(24.199_812_4, false, false, &Locale::default())
            .as_deref(),
        Some("24.20")
    );
    assert_eq!(
        XlsxNumberFormat::Builtin(999).display(1.0, false, false, &Locale::default()),
        None
    );
    assert_eq!(
        XlsxNumberFormat::Custom(" General ".to_owned())
            .display(123_456_789_012.0, false, false, &Locale::default())
            .as_deref(),
        Some("123456789012")
    );
    assert_eq!(
        formats[0]
            .display(123_456_789_012.0, false, true, &Locale::default())
            .as_deref(),
        Some("1.23457E11")
    );
    assert_eq!(
        formats[0]
            .display(1E-11, false, false, &Locale::default())
            .as_deref(),
        Some("0")
    );
    assert_eq!(
        formats[0]
            .display(-1E-11, false, true, &Locale::default())
            .as_deref(),
        Some("-1E-11")
    );

    let mut archive = archive(&[(
        "sheet.xml",
        r#"<worksheet><sheetData>
<row r="2"><c r="B2" s="1"><v><![CDATA[2087.0249999999996]]></v></c><c t="inlineStr"><is><t>text</t></is></c></row>
<row><c><v>0</v></c><c s="99"><v>1</v></c><c t="n"></c><c t="b"><v>true</v></c></row>
</sheetData></worksheet>"#,
    )]);
    let reader = display_xml_reader(&mut archive, "sheet.xml")?;
    let mut cache = create_cache(ReadCacheMode::Memory, 0)?;
    let mut cells = XlsxDisplayCellReader::new(
        reader,
        &formats,
        false,
        false,
        Locale::default(),
        cache.as_mut(),
    )?;

    let first = cells.next_cell()?.expect("first display cell");
    assert_eq!(first.position, (1, 1));
    assert_eq!(first.display_value.as_deref(), Some("2087.03"));
    assert_eq!(
        first.decimal_value.expect("decimal").to_string(),
        "2087.025"
    );
    let inline = cells.next_cell()?.expect("inline string cell");
    assert_eq!(inline.position, (1, 2));
    assert_eq!(inline.display_value, None);
    assert_eq!(inline.decimal_value, None);
    let zero = cells.next_cell()?.expect("zero cell");
    assert_eq!(zero.position, (2, 0));
    assert_eq!(zero.display_value.as_deref(), Some("0"));
    assert_eq!(zero.decimal_value.expect("zero decimal").to_string(), "0");
    let unknown_style = cells.next_cell()?.expect("unknown-style cell");
    assert_eq!(unknown_style.position, (2, 1));
    assert_eq!(unknown_style.display_value, None);
    assert_eq!(
        unknown_style.decimal_value.expect("decimal").to_string(),
        "1"
    );
    let empty_numeric = cells.next_cell()?.expect("empty numeric cell");
    assert_eq!(empty_numeric.position, (2, 2));
    assert_eq!(empty_numeric.display_value, None);
    assert_eq!(empty_numeric.decimal_value, None);
    let boolean = cells.next_cell()?.expect("boolean cell");
    assert_eq!(boolean.value, CellValue::Bool(true));
    assert!(cells.next_cell()?.is_none());
    let infinity = excel_display_number(f64::INFINITY);
    assert!(infinity.is_infinite() && infinity.is_sign_positive());
    Ok(())
}

#[test]
fn display_formats_numbers_dates_and_scientific_values_with_selected_locale() {
    let german = crate::ExcelLocale::from_name("de_DE")
        .expect("German locale")
        .formatter();
    assert_eq!(
        XlsxNumberFormat::Custom("#,##0.00".to_owned())
            .display(1_234.5, false, false, &german)
            .as_deref(),
        Some("1.234,50")
    );
    assert_eq!(
        XlsxNumberFormat::Builtin(0)
            .display(123_456_789_012.0, false, true, &german)
            .as_deref(),
        Some("1,23457E11")
    );

    let chinese = crate::ExcelLocale::from_name("zh_CN")
        .expect("Chinese locale")
        .formatter();
    assert_eq!(
        XlsxNumberFormat::Custom("mmmm dddd AM/PM".to_owned())
            .display(1.25, false, false, &chinese)
            .as_deref(),
        Some("一月 星期日 上午")
    );
}

#[test]
fn display_cell_stream_reports_every_xml_and_coordinate_boundary() -> Result<()> {
    let formats = vec![XlsxNumberFormat::Builtin(0)];
    for xml in ["<worksheet/>", "<worksheet><"] {
        let mut archive = archive(&[("sheet.xml", xml)]);
        let reader = display_xml_reader(&mut archive, "sheet.xml")?;
        let mut cache = create_cache(ReadCacheMode::Memory, 0)?;
        assert!(
            XlsxDisplayCellReader::new(
                reader,
                &formats,
                false,
                false,
                Locale::default(),
                cache.as_mut(),
            )
            .is_err()
        );
    }

    for xml in [
        "<worksheet><sheetData>",
        "<worksheet><sheetData><row r=></row></sheetData></worksheet>",
        "<worksheet><sheetData><row r=\"0\"></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c r=></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c r=\"XFE1\"></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c s=\"bad\"></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c><v>bad</v></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c><v>NaN</v></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c t=\"s\"><v>bad</v></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c t=\"s\"><v>999</v></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c t=\"unsupported\"><v>1</v></c></row></sheetData></worksheet>",
        "<worksheet><sheetData><row><c><v><",
        "<worksheet><sheetData><row><c><v>1",
        "<worksheet><sheetData><",
    ] {
        let mut archive = archive(&[("sheet.xml", xml)]);
        let reader = display_xml_reader(&mut archive, "sheet.xml")?;
        let mut cache = create_cache(ReadCacheMode::Memory, 0)?;
        let mut cells = XlsxDisplayCellReader::new(
            reader,
            &formats,
            false,
            false,
            Locale::default(),
            cache.as_mut(),
        )?;
        assert!(cells.next_cell().is_err(), "{xml}");
    }

    for bytes in [
        b"<worksheet><sheetData><row><c><v>\xff</v></c></row></sheetData></worksheet>".as_slice(),
        b"<worksheet><sheetData><row><c><v><![CDATA[\xff]]></v></c></row></sheetData></worksheet>"
            .as_slice(),
        b"<worksheet><sheetData><row><c><f>\xff</f></c></row></sheetData></worksheet>".as_slice(),
        b"<worksheet><sheetData><row><c><f><![CDATA[\xff]]></f></c></row></sheetData></worksheet>"
            .as_slice(),
        b"<worksheet><sheetData><row><c t=\"inlineStr\"><is><t>\xff</t></is></c></row></sheetData></worksheet>"
            .as_slice(),
        b"<worksheet><sheetData><row><c t=\"inlineStr\"><is><t><![CDATA[\xff]]></t></is></c></row></sheetData></worksheet>"
            .as_slice(),
    ] {
        let mut archive =
            ZipArchive::new(package_bytes(&[("sheet.xml", bytes)])).expect("ZIP archive");
        let reader = display_xml_reader(&mut archive, "sheet.xml")?;
        let mut cache = create_cache(ReadCacheMode::Memory, 0)?;
        let mut cells = XlsxDisplayCellReader::new(
            reader,
            &formats,
            false,
            false,
            Locale::default(),
            cache.as_mut(),
        )?;
        assert!(cells.next_cell().is_err());
    }
    Ok(())
}

struct RejectingSharedStringCache;

impl SharedStringCacheWriter for RejectingSharedStringCache {
    fn put(&mut self, _value: String) -> Result<()> {
        Err(ExcelError::Format(
            "injected shared-string cache failure".to_owned(),
        ))
    }

    fn finish(self: Box<Self>) -> Result<Box<dyn SharedStringCacheReader>> {
        Err(ExcelError::Format("unused shared-string finish".to_owned()))
    }
}

impl SharedStringCacheReader for RejectingSharedStringCache {
    fn get(&self, _index: usize) -> Result<String> {
        Err(ExcelError::Format("unused shared-string lookup".to_owned()))
    }

    fn len(&self) -> usize {
        0
    }
}

impl SharedStringCache for RejectingSharedStringCache {}

#[test]
fn shared_string_stream_propagates_package_xml_utf8_factory_and_cache_failures() {
    let mut missing = archive(&[]);
    let cache = path_cache(&missing);
    assert!(
        read_shared_strings(&mut missing, &cache, "missing.xml", ReadCacheMode::Memory).is_err()
    );

    let mut malformed = archive(&[("shared.xml", "<sst><si><")]);
    let cache = path_cache(&malformed);
    assert!(
        read_shared_strings(&mut malformed, &cache, "shared.xml", ReadCacheMode::Memory).is_err()
    );

    for bytes in [
        b"<sst><si><t>\xff</t></si></sst>".as_slice(),
        b"<sst><si><t><![CDATA[\xff]]></t></si></sst>".as_slice(),
    ] {
        let mut archive = ZipArchive::new(package_bytes(&[("shared.xml", bytes)]))
            .expect("shared-string ZIP archive");
        let cache = path_cache(&archive);
        assert!(
            read_shared_strings(&mut archive, &cache, "shared.xml", ReadCacheMode::Memory).is_err()
        );
    }

    let mut archive = archive(&[("shared.xml", "<sst><si><t>value</t></si></sst>")]);
    let cache = path_cache(&archive);
    let mut failing_factory = |_| Err(ExcelError::Format("injected factory failure".to_owned()));
    assert!(
        read_shared_strings_with_factory(&mut archive, &cache, "shared.xml", &mut failing_factory,)
            .is_err()
    );

    let mut rejecting = RejectingSharedStringCache;
    let mut input = Cursor::new(b"<sst><si><t>value</t></si></sst>".as_slice());
    assert!(parse_shared_strings(&mut input, &mut rejecting).is_err());
    assert!(rejecting.get(0).is_err());
    assert_eq!(rejecting.len(), 0);
}

#[test]
fn style_parser_maps_custom_builtin_defaults_and_rejects_invalid_xml() -> Result<()> {
    let xml = r#"<styleSheet>
<numFmts><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165"/></numFmts>
<cellXfs><xf numFmtId="164"/><xf numFmtId="14"/><xf/></cellXfs>
</styleSheet>"#;
    let mut styles = archive(&[("styles.xml", xml)]);
    let cache = path_cache(&styles);
    let expected = vec![
        XlsxNumberFormat::Custom("0.00".to_owned()),
        XlsxNumberFormat::Builtin(14),
        XlsxNumberFormat::Builtin(0),
    ];
    assert_eq!(
        read_cell_formats(&mut styles, &cache, "styles.xml")?,
        expected
    );
    let mut styles = xlsx_input_archive(&[("styles.xml", xml)]);
    let cache = path_cache(&styles);
    assert_eq!(
        read_cell_formats(&mut styles, &cache, "styles.xml")?,
        expected
    );

    let mut empty = archive(&[("styles.xml", "<styleSheet/>")]);
    let cache = path_cache(&empty);
    assert_eq!(
        read_cell_formats(&mut empty, &cache, "styles.xml")?,
        vec![XlsxNumberFormat::Builtin(0)]
    );
    let mut empty = xlsx_input_archive(&[("styles.xml", "<styleSheet/>")]);
    let cache = path_cache(&empty);
    assert_eq!(
        read_cell_formats(&mut empty, &cache, "styles.xml")?,
        vec![XlsxNumberFormat::Builtin(0)]
    );

    let mut missing = archive(&[]);
    let cache = path_cache(&missing);
    assert!(read_cell_formats(&mut missing, &cache, "missing.xml").is_err());
    let mut missing = xlsx_input_archive(&[]);
    let cache = path_cache(&missing);
    assert!(read_cell_formats(&mut missing, &cache, "missing.xml").is_err());
    for xml in [
        "<styleSheet>",
        "<styleSheet><",
        "<styleSheet><numFmt numFmtId=\"bad\" formatCode=\"0\"/></styleSheet>",
        "<styleSheet><cellXfs><xf numFmtId=\"bad\"/></cellXfs></styleSheet>",
        "<styleSheet><cellXfs><xf numFmtId=></xf></cellXfs></styleSheet>",
        "<styleSheet><numFmt numFmtId=></numFmt></styleSheet>",
    ] {
        let mut styles = archive(&[("styles.xml", xml)]);
        let cache = path_cache(&styles);
        assert!(read_cell_formats(&mut styles, &cache, "styles.xml").is_err());
        let mut styles = xlsx_input_archive(&[("styles.xml", xml)]);
        let cache = path_cache(&styles);
        assert!(read_cell_formats(&mut styles, &cache, "styles.xml").is_err());
    }
    Ok(())
}

#[test]
fn row_scanner_handles_empty_implicit_limits_and_malformed_xml() -> Result<()> {
    assert_eq!(
        scan_last_row(Cursor::new(
            b"<worksheet><sheetData/></worksheet>".as_slice(),
        ))?,
        None
    );
    assert_eq!(
        scan_last_row(Cursor::new(
            br#"<worksheet><sheetData><row r="2"/><row/><row r="1048576"/></sheetData></worksheet>"#.as_slice(),
        ))?,
        Some(MAX_XLSX_ROW_NUMBER - 1)
    );
    assert!(scan_last_row(Cursor::new(b"<worksheet>".as_slice())).is_err());
    assert!(
        scan_last_row(Cursor::new(
            b"<worksheet><sheetData><row r=></row></sheetData></worksheet>".as_slice()
        ))
        .is_err()
    );
    for malformed in [
        b"<worksheet><sheetData><".as_slice(),
        b"<worksheet><sheetData><row r=\"0\"/></sheetData></worksheet>".as_slice(),
        b"<worksheet><sheetData><row \xff=\"x\"/></sheetData></worksheet>".as_slice(),
        b"<worksheet><sheetData><row r=\"\xff\"/></sheetData></worksheet>".as_slice(),
    ] {
        assert!(scan_last_row(Cursor::new(malformed)).is_err());
    }
    assert_eq!(parse_row_number("1")?, 0);
    for invalid in ["not-a-row", "0", "1048577"] {
        assert!(parse_row_number(invalid).is_err());
    }
    Ok(())
}

#[test]
fn relationship_paths_normalize_absolute_relative_and_invalid_targets() -> Result<()> {
    assert_eq!(
        relationship_part_name("workbook.xml"),
        "_rels/workbook.xml.rels"
    );
    assert_eq!(
        relationship_part_name("xl/workbook.xml"),
        "xl/_rels/workbook.xml.rels"
    );
    assert_eq!(resolve_target("", "xl/workbook.xml")?, "xl/workbook.xml");
    assert_eq!(
        resolve_target("xl/workbook.xml", "worksheets/./sheet1.xml")?,
        "xl/worksheets/sheet1.xml"
    );
    assert_eq!(
        resolve_target("xl/workbook.xml", "/xl/worksheets/sheet1.xml")?,
        "xl/worksheets/sheet1.xml"
    );
    assert_eq!(
        resolve_target("xl/workbook.xml", "../workbook.xml")?,
        "workbook.xml"
    );
    assert!(resolve_target("", "../outside.xml").is_err());
    assert!(resolve_target("", "./").is_err());

    let mut cache = HashMap::new();
    cache.insert("xl/workbook.xml".to_owned(), "XL/Workbook.xml".to_owned());
    assert_eq!(cached_path(&cache, "xl/workbook.xml"), "XL/Workbook.xml");
    assert_eq!(cached_path(&cache, "missing.xml"), "missing.xml");
    Ok(())
}

#[test]
fn relationship_parser_rejects_missing_truncated_and_malformed_parts() {
    let mut missing = archive(&[]);
    let cache = path_cache(&missing);
    assert!(read_relationships(&mut missing, &cache, "missing.rels").is_err());

    let mut truncated = archive(&[("rels.xml", "<Relationships>")]);
    let cache = path_cache(&truncated);
    assert!(read_relationships(&mut truncated, &cache, "rels.xml").is_err());

    let mut malformed = archive(&[(
        "rels.xml",
        "<Relationships><Relationship Id=></Relationship></Relationships>",
    )]);
    let cache = path_cache(&malformed);
    assert!(read_relationships(&mut malformed, &cache, "rels.xml").is_err());

    let mut malformed_event = archive(&[("rels.xml", "<Relationships><")]);
    let cache = path_cache(&malformed_event);
    assert!(read_relationships(&mut malformed_event, &cache, "rels.xml").is_err());

    let mut missing = xlsx_input_archive(&[]);
    let cache = path_cache(&missing);
    assert!(read_raw_relationships(&mut missing, &cache, "missing.rels").is_err());
    for xml in [
        "<Relationships>",
        "<Relationships><Relationship Id=></Relationship></Relationships>",
        "<Relationships><",
    ] {
        let mut invalid = xlsx_input_archive(&[("rels.xml", xml)]);
        let cache = path_cache(&invalid);
        assert!(read_raw_relationships(&mut invalid, &cache, "rels.xml").is_err());
    }
    let mut missing_id = xlsx_input_archive(&[(
        "rels.xml",
        "<Relationships><Relationship Target=\"ignored\"/></Relationships>",
    )]);
    let cache = path_cache(&missing_id);
    assert!(
        read_raw_relationships(&mut missing_id, &cache, "rels.xml")
            .expect("relationship table")
            .is_empty()
    );
}

#[test]
fn workbook_parser_rejects_missing_attributes_relationships_and_xml() {
    let worksheet_relationships = HashMap::from([(
        "sheet".to_owned(),
        (
            "worksheets/sheet1.xml".to_owned(),
            "urn:relationships/worksheet".to_owned(),
        ),
    )]);
    for xml in [
        "<workbook><sheets><sheet id=\"sheet\"/></sheets></workbook>",
        "<workbook><sheets><sheet name=\"Name\"/></sheets></workbook>",
        "<workbook><sheets><sheet name=\"Name\" id=\"missing\"/></sheets></workbook>",
        "<workbook><sheets><sheet name=\"Name\" id=></sheet></sheets></workbook>",
        "<workbook><workbookPr date1904=></workbookPr></workbook>",
        "<workbook>",
        "<workbook><",
    ] {
        let mut workbook = archive(&[("xl/workbook.xml", xml)]);
        let cache = path_cache(&workbook);
        assert!(
            read_workbook_metadata(
                &mut workbook,
                &cache,
                "xl/workbook.xml",
                &worksheet_relationships,
            )
            .is_err()
        );
    }

    let mut missing = archive(&[]);
    let cache = path_cache(&missing);
    assert!(
        read_workbook_metadata(
            &mut missing,
            &cache,
            "xl/workbook.xml",
            &worksheet_relationships,
        )
        .is_err()
    );

    let escaping = HashMap::from([(
        "sheet".to_owned(),
        (
            "../../outside.xml".to_owned(),
            "urn:relationships/worksheet".to_owned(),
        ),
    )]);
    let mut workbook = archive(&[(
        "xl/workbook.xml",
        "<workbook><sheets><sheet name=\"Name\" id=\"sheet\"/></sheets></workbook>",
    )]);
    let cache = path_cache(&workbook);
    assert!(read_workbook_metadata(&mut workbook, &cache, "xl/workbook.xml", &escaping,).is_err());
}

#[test]
fn metadata_constructor_reports_each_package_boundary() {
    assert!(XlsxRowMetadata::new(Cursor::new(b"not-a-zip".to_vec())).is_err());
    assert!(XlsxRowMetadata::new(package(&[])).is_err());
    assert!(XlsxRowMetadata::new(package(&[("_rels/.rels", "<Relationships/>"),])).is_err());
    assert!(
        XlsxRowMetadata::new(package(&[(
            "_rels/.rels",
            "<Relationships><Relationship Id=\"office\" Type=\"x/officeDocument\" Target=\"../outside.xml\"/></Relationships>",
        )]))
        .is_err()
    );
    assert!(
        XlsxRowMetadata::new(package(&[(
            "_rels/.rels",
            "<Relationships><Relationship Id=\"office\" Type=\"x/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
        )]))
        .is_err()
    );
    assert!(
        XlsxRowMetadata::new(package(&[
            (
                "_rels/.rels",
                "<Relationships><Relationship Id=\"office\" Type=\"x/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>",
            ),
            (
                "xl/_rels/workbook.xml.rels",
                "<Relationships><Relationship Id=\"sheet\" Type=\"x/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>",
            ),
            ("xl/workbook.xml", "<workbook>"),
        ]))
        .is_err()
    );

    let base_relationships = "<Relationships><Relationship Id=\"sheet\" Type=\"x/worksheet\" Target=\"sheets/First.XML\"/>";
    for (styles_target, styles_xml) in [
        ("../../../outside.xml", None),
        ("styles.xml", Some("<styleSheet>")),
    ] {
        let workbook_relationships = format!(
            "{base_relationships}<Relationship Id=\"styles\" Type=\"x/styles\" Target=\"{styles_target}\"/></Relationships>"
        );
        let workbook = "<workbook><sheets><sheet name=\"A\" id=\"sheet\"/></sheets></workbook>";
        let mut entries = vec![
            (
                "_rels/.rels",
                "<Relationships><Relationship Id=\"office\" Type=\"x/officeDocument\" Target=\"custom/workbook.xml\"/></Relationships>",
            ),
            ("custom/workbook.xml", workbook),
            (
                "custom/sheets/First.XML",
                "<worksheet><sheetData/></worksheet>",
            ),
            (
                "custom/_rels/workbook.xml.rels",
                workbook_relationships.as_str(),
            ),
        ];
        if let Some(styles_xml) = styles_xml {
            entries.push(("custom/styles.xml", styles_xml));
        }
        assert!(XlsxRowMetadata::new(package(&entries)).is_err());
        assert!(XlsxRowMetadata::new(xlsx_input(&entries)).is_err());
    }

    for shared_target in ["../../../outside.xml", "missing.xml"] {
        let workbook_relationships = format!(
            "{base_relationships}<Relationship Id=\"shared\" Type=\"x/sharedStrings\" Target=\"{shared_target}\"/></Relationships>"
        );
        let entries = [
            (
                "_rels/.rels",
                "<Relationships><Relationship Id=\"office\" Type=\"x/officeDocument\" Target=\"custom/workbook.xml\"/></Relationships>",
            ),
            (
                "custom/workbook.xml",
                "<workbook><sheets><sheet name=\"A\" id=\"sheet\"/></sheets></workbook>",
            ),
            (
                "custom/sheets/First.XML",
                "<worksheet><sheetData/></worksheet>",
            ),
            (
                "custom/_rels/workbook.xml.rels",
                workbook_relationships.as_str(),
            ),
        ];
        assert!(XlsxRowMetadata::new(package(&entries)).is_err());
        assert!(XlsxRowMetadata::new(xlsx_input(&entries)).is_err());
    }
}

#[test]
fn extras_parse_merge_internal_external_hyperlinks_and_rich_comments() -> Result<()> {
    let mut entries = valid_entries();
    entries.pop();
    entries.extend([
        (
            "custom/sheets/First.XML",
            r#"<worksheet xmlns:r="urn:r"><sheetData><row r="1"/></sheetData>
<mergeCells><mergeCell ref="$A$1:B2"/></mergeCells>
<hyperlinks><hyperlink ref="C3" location="'Other Sheet'!A1"/><hyperlink ref="D4:E5" r:id="external"/></hyperlinks>
</worksheet>"#,
        ),
        (
            "custom/sheets/_rels/First.XML.rels",
            r#"<Relationships>
<Relationship Id="external" Type="urn:relationships/hyperlink" Target="https://example.com?a=1&amp;b=2" TargetMode="External"/>
<Relationship Id="comments" Type="urn:relationships/comments" Target="../comments1.xml"/>
</Relationships>"#,
        ),
        (
            "custom/comments1.xml",
            r#"<comments><commentList><comment ref="F6"><text><r><t>A&amp;B</t></r><r><t> plus </t></r><r><t><![CDATA[C]]></t></r><r><t>&#x21;</t></r></text></comment></commentList></comments>"#,
        ),
    ]);
    let mut metadata = XlsxRowMetadata::new(package(&entries))?;
    let enabled = HashSet::from([
        easyexcel_core::CellExtraType::Comment,
        easyexcel_core::CellExtraType::Hyperlink,
        easyexcel_core::CellExtraType::Merge,
    ]);
    let extras = metadata.extras("A&B", &enabled)?;
    assert_eq!(extras.len(), 4);
    assert_eq!(extras[0].extra_type(), easyexcel_core::CellExtraType::Merge);
    assert_eq!(extras[0].text(), None);
    assert_eq!(extras[0].first_row_index(), 0);
    assert_eq!(extras[0].last_row_index(), 1);
    assert_eq!(extras[0].first_column_index(), 0);
    assert_eq!(extras[0].last_column_index(), 1);
    assert_eq!(extras[1].extra_type(), easyexcel_core::CellExtraType::Hyperlink);
    assert_eq!(extras[1].text(), Some("'Other Sheet'!A1"));
    assert_eq!(extras[2].text(), Some("https://example.com?a=1&b=2"));
    assert_eq!(extras[2].first_row_index(), 3);
    assert_eq!(extras[2].last_row_index(), 4);
    assert_eq!(extras[2].first_column_index(), 3);
    assert_eq!(extras[2].last_column_index(), 4);
    assert_eq!(extras[3].extra_type(), easyexcel_core::CellExtraType::Comment);
    assert_eq!(extras[3].text(), Some("A&B plus C!"));
    assert_eq!(extras[3].first_row_index(), 5);
    assert_eq!(extras[3].last_row_index(), 5);
    assert_eq!(extras[3].first_column_index(), 5);
    assert_eq!(extras[3].last_column_index(), 5);

    assert!(metadata.extras("Missing", &enabled).is_err());
    assert!(metadata.extras("A&B", &HashSet::new())?.is_empty());
    let mut metadata = XlsxRowMetadata::new(xlsx_input(&entries))?;
    assert!(metadata.extras("Missing", &enabled).is_err());
    Ok(())
}

#[test]
fn cell_reference_parser_enforces_xlsx_coordinates_and_range_ordering() -> Result<()> {
    assert_eq!(parse_cell_reference("A1")?, (0, 0));
    assert_eq!(parse_cell_reference("$xFd$1048576")?, (1_048_575, 16_383));
    assert_eq!(parse_cell_range("B2")?, (1, 1, 1, 1));
    assert_eq!(parse_cell_range("A1:B2")?, (0, 1, 0, 1));
    for invalid in [
        "",
        "1",
        "A",
        "A0",
        "XFE1",
        "A1048577",
        "A-1",
        "A1:B",
        "B2:A1",
        "A1:B2:C3",
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1",
    ] {
        assert!(parse_cell_range(invalid).is_err(), "{invalid}");
    }
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn extra_parsers_report_missing_relationships_attributes_parts_and_xml() {
    let enabled = HashSet::from([easyexcel_core::CellExtraType::Hyperlink, easyexcel_core::CellExtraType::Merge]);
    let relationships = RawRelationships::new();
    for xml in [
        "<worksheet><mergeCell/></worksheet>",
        "<worksheet><mergeCell ref=\"XFE1\"/></worksheet>",
        "<worksheet><hyperlink/></worksheet>",
        "<worksheet><hyperlink ref=\"A1\"/></worksheet>",
        "<worksheet><hyperlink ref=\"XFE1\" location=\"Sheet2!A1\"/></worksheet>",
        "<worksheet><hyperlink ref=></hyperlink></worksheet>",
        "<worksheet><hyperlink ref=\"A1\" id=\"missing\"/></worksheet>",
        "<worksheet>",
        "<worksheet><",
    ] {
        let mut archive = archive(&[("sheet.xml", xml)]);
        let cache = path_cache(&archive);
        assert!(
            read_worksheet_extras(&mut archive, &cache, "sheet.xml", &relationships, &enabled)
                .is_err()
        );
    }

    let mut missing = archive(&[]);
    let cache = path_cache(&missing);
    assert!(
        read_worksheet_extras(
            &mut missing,
            &cache,
            "missing.xml",
            &relationships,
            &enabled
        )
        .is_err()
    );

    let mut malformed_attribute = ZipArchive::new(package_bytes(&[(
        "sheet.xml",
        b"<worksheet><mergeCell \xff=\"x\"/></worksheet>",
    )]))
    .expect("ZIP archive");
    let cache = path_cache(&malformed_attribute);
    assert!(
        read_worksheet_extras(
            &mut malformed_attribute,
            &cache,
            "sheet.xml",
            &relationships,
            &enabled
        )
        .is_err()
    );

    for xml in [
        "<comments><commentList><comment><text><t>x</t></text></comment></commentList></comments>",
        "<comments><commentList><comment ref=\"XFE1\"><text><t>x</t></text></comment></commentList></comments>",
        "<comments><commentList></comment></commentList></comments>",
        "<comments><commentList><comment ref=\"A1\"><text><t>&unknown;</t></text></comment></commentList></comments>",
        "<comments><commentList><comment ref=\"A1\"><text><t>&#0;</t></text></comment></commentList></comments>",
        "<comments>",
        "<comments><",
        "<comments><commentList><comment ref=\"A1\"><text><t>\u{fffd}</t></text></comment></commentList></comments>",
    ] {
        let mut archive = archive(&[("comments.xml", xml)]);
        let cache = path_cache(&archive);
        let result = read_comments(&mut archive, &cache, "comments.xml");
        if xml.contains('\u{fffd}') {
            assert!(result.is_ok());
        } else {
            assert!(result.is_err());
        }
    }
    let mut missing_comments = archive(&[]);
    let cache = path_cache(&missing_comments);
    assert!(read_comments(&mut missing_comments, &cache, "missing.xml").is_err());

    let mut missing_xlsx_part = xlsx_input_archive(&[]);
    let cache = path_cache(&missing_xlsx_part);
    assert!(
        read_worksheet_extras(
            &mut missing_xlsx_part,
            &cache,
            "missing-sheet.xml",
            &relationships,
            &enabled,
        )
        .is_err()
    );
    assert!(read_comments(&mut missing_xlsx_part, &cache, "missing-comments.xml").is_err());

    for bytes in [
        b"<comments><commentList><comment \xff=\"x\"></comment></commentList></comments>".as_slice(),
        b"<comments><commentList><comment ref=\"A1\"><text><t>\xff</t></text></comment></commentList></comments>".as_slice(),
        b"<comments><commentList><comment ref=\"A1\"><text><t><![CDATA[\xff]]></t></text></comment></commentList></comments>".as_slice(),
        b"<comments><commentList><comment ref=\"A1\"><text><t>&\xff;</t></text></comment></commentList></comments>".as_slice(),
    ] {
        let mut invalid = ZipArchive::new(package_bytes(&[("comments.xml", bytes)]))
            .expect("ZIP archive");
        let cache = path_cache(&invalid);
        assert!(read_comments(&mut invalid, &cache, "comments.xml").is_err());
    }

    let mut entries = valid_entries();
    entries.pop();
    entries.extend([
        (
            "custom/sheets/First.XML",
            "<worksheet><sheetData/></worksheet>",
        ),
        (
            "custom/sheets/_rels/First.XML.rels",
            "<Relationships><Relationship Id=\"comments\" Type=\"urn:relationships/comments\" Target=\"../missing-comments.xml\"/></Relationships>",
        ),
    ]);
    let mut metadata = XlsxRowMetadata::new(package(&entries)).expect("metadata package");
    assert!(
        metadata
            .extras("A&B", &HashSet::from([easyexcel_core::CellExtraType::Comment]))
            .is_err()
    );
    let mut metadata = XlsxRowMetadata::new(xlsx_input(&entries)).expect("metadata package");
    assert!(
        metadata
            .extras("A&B", &HashSet::from([easyexcel_core::CellExtraType::Comment]))
            .is_err()
    );

    let mut invalid_relationships = valid_entries();
    invalid_relationships.pop();
    invalid_relationships.extend([
        (
            "custom/sheets/First.XML",
            "<worksheet><sheetData/></worksheet>",
        ),
        ("custom/sheets/_rels/First.XML.rels", "<Relationships>"),
    ]);
    let mut metadata =
        XlsxRowMetadata::new(package(&invalid_relationships)).expect("metadata package");
    assert!(
        metadata
            .extras("A&B", &HashSet::from([easyexcel_core::CellExtraType::Hyperlink]))
            .is_err()
    );
    let mut metadata =
        XlsxRowMetadata::new(xlsx_input(&invalid_relationships)).expect("metadata package");
    assert!(
        metadata
            .extras("A&B", &HashSet::from([easyexcel_core::CellExtraType::Hyperlink]))
            .is_err()
    );

    let mut escaping_comment = valid_entries();
    escaping_comment.pop();
    escaping_comment.extend([
        (
            "custom/sheets/First.XML",
            "<worksheet><sheetData/></worksheet>",
        ),
        (
            "custom/sheets/_rels/First.XML.rels",
            "<Relationships><Relationship Id=\"comments\" Type=\"urn:relationships/comments\" Target=\"../../../../outside.xml\"/></Relationships>",
        ),
    ]);
    let mut metadata = XlsxRowMetadata::new(package(&escaping_comment)).expect("metadata package");
    assert!(
        metadata
            .extras("A&B", &HashSet::from([easyexcel_core::CellExtraType::Comment]))
            .is_err()
    );
    let mut metadata =
        XlsxRowMetadata::new(xlsx_input(&escaping_comment)).expect("metadata package");
    assert!(
        metadata
            .extras("A&B", &HashSet::from([easyexcel_core::CellExtraType::Comment]))
            .is_err()
    );
}
