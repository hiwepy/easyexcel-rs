use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use super::*;

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

fn archive(entries: &[(&str, &str)]) -> ZipArchive<Cursor<Vec<u8>>> {
    ZipArchive::new(package(entries)).expect("valid ZIP archive")
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
    assert_eq!(metadata.last_explicit_row("A&B")?, Some(4));
    assert!(metadata.last_explicit_row("Chart").is_err());
    assert!(metadata.last_explicit_row("Missing").is_err());

    let mut missing_part = XlsxRowMetadata {
        archive: archive(&[]),
        path_cache: HashMap::new(),
        sheet_paths: HashMap::from([("MissingPart".to_owned(), "missing.xml".to_owned())]),
    };
    assert!(missing_part.last_explicit_row("MissingPart").is_err());

    let mut missing_sheet_entries = valid_entries();
    missing_sheet_entries.pop();
    assert!(XlsxRowMetadata::new(package(&missing_sheet_entries)).is_err());
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
        "<workbook>",
        "<workbook><",
    ] {
        let mut workbook = archive(&[("xl/workbook.xml", xml)]);
        let cache = path_cache(&workbook);
        assert!(
            read_sheet_paths(
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
        read_sheet_paths(
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
    assert!(read_sheet_paths(&mut workbook, &cache, "xl/workbook.xml", &escaping,).is_err());
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
}
