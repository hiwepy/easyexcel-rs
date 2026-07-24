//! Java `@NumberFormat` converter and derive integration coverage.

use easyexcel::{EasyExcel, ExcelRow, NumberRoundingMode, Result};
use std::io::Write;
use tempfile::tempdir;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct PercentageRow {
    #[excel(
        name = "percentage",
        index = 0,
        number_format = "#.##%",
        rounding_mode = "HALF_UP"
    )]
    value: f64,
}

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct HalfDownRow {
    #[excel(
        name = "amount",
        index = 0,
        number_format = "0.00",
        rounding_mode = "HALF_DOWN"
    )]
    value: f64,
}

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct StringCellRow {
    #[excel(name = "percentage", index = 0)]
    percentage: String,
    #[excel(name = "error", index = 1)]
    error: String,
}

fn write_string_converter_fixture(path: &std::path::Path) -> Result<()> {
    let mut archive = ZipWriter::new(std::fs::File::create(path)?);
    let entries = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
        ),
        (
            "xl/workbook.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
        ),
        (
            "xl/styles.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>
  <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs>
</styleSheet>"#,
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>percentage</t></is></c>
      <c r="B1" t="inlineStr"><is><t>error</t></is></c>
    </row>
    <row r="2">
      <c r="A2" s="1"><v>0.1250</v></c>
      <c r="B2" t="e"><v>#DIV/0!</v></c>
    </row>
  </sheetData>
</worksheet>"#,
        ),
    ];
    for (name, contents) in entries {
        archive
            .start_file(name, SimpleFileOptions::default())
            .map_err(|error| easyexcel::ExcelError::Format(error.to_string()))?;
        archive.write_all(contents.as_bytes())?;
    }
    archive
        .finish()
        .map_err(|error| easyexcel::ExcelError::Format(error.to_string()))?;
    Ok(())
}

#[test]
fn derive_number_format_round_trips_through_default_csv_string_converters() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("percentage.csv");
    let rows = vec![
        PercentageRow { value: 1.235 },
        PercentageRow { value: -0.125 },
    ];

    EasyExcel::write::<PercentageRow>(&path)
        .with_bom(false)
        .do_write(rows.clone())?;

    let contents = std::fs::read_to_string(&path)?;
    assert!(contents.contains("123.5%"));
    assert!(contents.contains("-12.5%"));
    assert_eq!(
        EasyExcel::read_sync::<PercentageRow>(&path).do_read_sync()?,
        rows
    );
    assert_eq!(
        PercentageRow::schema()[0].number_rounding_mode,
        Some(NumberRoundingMode::HalfUp)
    );
    Ok(())
}

#[test]
fn derive_half_down_rounding_is_used_by_csv_writer() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("half-down.csv");
    EasyExcel::write::<HalfDownRow>(&path)
        .with_bom(false)
        .do_write([HalfDownRow { value: 1.225 }])?;

    let contents = std::fs::read_to_string(path)?;
    assert!(contents.contains("1.22"));
    assert_eq!(
        HalfDownRow::schema()[0].number_rounding_mode,
        Some(NumberRoundingMode::HalfDown)
    );
    Ok(())
}

#[test]
fn xlsx_model_mapping_uses_source_display_and_error_converters_end_to_end() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("string-converters.xlsx");
    write_string_converter_fixture(&path)?;

    assert_eq!(
        EasyExcel::read_sync::<StringCellRow>(&path).do_read_sync()?,
        vec![StringCellRow {
            percentage: "12.50%".to_owned(),
            error: "#DIV/0!".to_owned(),
        }]
    );
    Ok(())
}
