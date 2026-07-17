use std::fs;
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};

use base64::Engine;
use calamine::{Data, Reader, Xlsx, open_workbook};
use flate2::read::GzDecoder;
use rust_xlsxwriter::{Format, Workbook};
use tempfile::{TempDir, tempdir};

use super::*;

struct FaultyIo {
    inner: Cursor<Vec<u8>>,
    fail_read_at: Option<usize>,
    fail_write_at: Option<usize>,
    fail_seek_at: Option<usize>,
    reads: usize,
    writes: usize,
    seeks: usize,
}

impl FaultyIo {
    fn reading(bytes: Vec<u8>, fail_at: usize) -> Self {
        Self {
            inner: Cursor::new(bytes),
            fail_read_at: Some(fail_at),
            fail_write_at: None,
            fail_seek_at: None,
            reads: 0,
            writes: 0,
            seeks: 0,
        }
    }

    fn writing(fail_at: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            fail_read_at: None,
            fail_write_at: Some(fail_at),
            fail_seek_at: None,
            reads: 0,
            writes: 0,
            seeks: 0,
        }
    }

    fn seeking(bytes: Vec<u8>, fail_at: usize) -> Self {
        Self {
            inner: Cursor::new(bytes),
            fail_read_at: None,
            fail_write_at: None,
            fail_seek_at: Some(fail_at),
            reads: 0,
            writes: 0,
            seeks: 0,
        }
    }
}

impl Read for FaultyIo {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let call = self.reads;
        self.reads += 1;
        if self.fail_read_at == Some(call) {
            return Err(io::Error::other("injected read failure"));
        }
        self.inner.read(buffer)
    }
}

impl Write for FaultyIo {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let call = self.writes;
        self.writes += 1;
        if self.fail_write_at == Some(call) {
            return Err(io::Error::other("injected write failure"));
        }
        self.inner.write(buffer)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for FaultyIo {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let call = self.seeks;
        self.seeks += 1;
        if self.fail_seek_at == Some(call) {
            return Err(io::Error::other("injected seek failure"));
        }
        self.inner.seek(position)
    }
}

fn test_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

fn successful_zip_operation(writer: &mut ArchiveWriter) -> Result<()> {
    writer.flush().map_err(ExcelError::from)
}

fn template_fixture() -> Result<(TempDir, std::path::PathBuf)> {
    let directory = tempdir()?;
    let path = directory.path().join("template.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Hello {name}")
        .map_err(test_error)?;
    worksheet
        .write_string(1, 0, "Count: {count}")
        .map_err(test_error)?;
    worksheet
        .write_string(2, 0, "Unknown: {unknown}")
        .map_err(test_error)?;
    workbook.save(&path).map_err(test_error)?;
    Ok((directory, path))
}

fn write_compressed_java_fixture(path: &Path, fixture: &str) -> Result<()> {
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(fixture.trim())
        .map_err(test_error)?;
    let mut decoder = GzDecoder::new(compressed.as_slice());
    let mut workbook = Vec::new();
    decoder.read_to_end(&mut workbook)?;
    fs::write(path, workbook)?;
    Ok(())
}

fn write_java_composite_fixture(path: &Path) -> Result<()> {
    write_compressed_java_fixture(
        path,
        include_str!("fixtures/java-demo-composite.xlsx.gz.b64"),
    )
}

fn find_string_coordinate(range: &calamine::Range<Data>, needle: &str) -> Option<(u32, u32)> {
    range.cells().find_map(|(row, column, value)| {
        (value == &Data::String(needle.to_owned())).then(|| {
            (
                u32::try_from(row).expect("small row"),
                u32::try_from(column).expect("small column"),
            )
        })
    })
}

#[test]
fn template_data_and_xml_escaping_are_deterministic() {
    let mut data = TemplateData::new().with("name", "Alice").with("count", 2);
    assert_eq!(
        data.insert("name", "Bob"),
        Some(CellValue::String("Alice".to_owned()))
    );
    assert_eq!(data.insert("new", "value"), None);
    assert_eq!(
        data.values().get("name"),
        Some(&CellValue::String("Bob".to_owned()))
    );
    assert_eq!(escape_xml("<&>\"' text"), "&lt;&amp;&gt;&quot;&apos; text");
    assert_eq!(
        replace_placeholders(
            "{a}-{missing}-{b}",
            &BTreeMap::from([
                ("a".to_owned(), CellValue::String("<".to_owned())),
                ("b".to_owned(), CellValue::String("&".to_owned()))
            ])
        ),
        "&lt;-{missing}-&amp;"
    );
    assert_eq!(
        replace_placeholders(
            r"\{a\}-{a}-\{missing\}",
            &BTreeMap::from([("a".to_owned(), CellValue::String("<值>".to_owned()))])
        ),
        "{a}-&lt;值&gt;-{missing}"
    );
    assert!(!contains_unescaped(r"\{users.name}", "{users."));
    assert!(contains_unescaped("{users.name}", "{users."));
    assert_eq!(TemplateData::default(), TemplateData::new());

    let owned = "owned".to_owned();
    let typed_date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    let date_time = typed_date.and_hms_opt(12, 34, 56).expect("valid time");
    for value in [
        "text".into_template_value(),
        owned.clone().into_template_value(),
        (&owned).into_template_value(),
        true.into_template_value(),
        i8::MIN.into_template_value(),
        i16::MIN.into_template_value(),
        i32::MIN.into_template_value(),
        i64::MIN.into_template_value(),
        isize::MIN.into_template_value(),
        i128::MIN.into_template_value(),
        u8::MAX.into_template_value(),
        u16::MAX.into_template_value(),
        u32::MAX.into_template_value(),
        usize::MAX.into_template_value(),
        u64::MAX.into_template_value(),
        u128::MAX.into_template_value(),
        BigInt::from(i128::MAX).into_template_value(),
        1.25_f32.into_template_value(),
        2.5_f64.into_template_value(),
        BigDecimal::from(42).into_template_value(),
        typed_date.into_template_value(),
        date_time.into_template_value(),
        Some(7_i32).into_template_value(),
        Option::<i32>::None.into_template_value(),
        CellValue::Error("#N/A".to_owned()).into_template_value(),
    ] {
        assert!(matches!(value, CellValue::Empty) || !value.as_text().is_empty());
    }
}

#[test]
fn exact_placeholders_preserve_java_scalar_cell_types() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("typed-template.xlsx");
    let output = directory.path().join("typed-output.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    for (row, placeholder) in [
        "{string}",
        "{boolean}",
        "{integer}",
        "{float}",
        "{decimal}",
        "{date}",
        "{datetime}",
        "{error}",
        "{formula}",
        "{empty}",
        "value={integer}",
    ]
    .into_iter()
    .enumerate()
    {
        worksheet
            .write_string(u32::try_from(row).expect("small row"), 0, placeholder)
            .map_err(test_error)?;
    }
    workbook.save(&template).map_err(test_error)?;

    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    fill_xlsx_template(
        &template,
        &output,
        &TemplateData::new()
            .with("string", "Alice")
            .with("boolean", true)
            .with("integer", 42_i64)
            .with("float", 5.25_f64)
            .with("decimal", BigDecimal::from(12345))
            .with("date", date)
            .with(
                "datetime",
                date.and_hms_opt(13, 14, 15).expect("valid time"),
            )
            .with("error", CellValue::Error("#N/A".to_owned()))
            .with("formula", CellValue::Formula("SUM(20,22)".to_owned()))
            .with("empty", Option::<String>::None),
    )?;

    let mut workbook: Xlsx<_> = open_workbook(&output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get_value((0, 0)),
        Some(&Data::String("Alice".to_owned()))
    );
    assert_eq!(range.get_value((1, 0)), Some(&Data::Bool(true)));
    assert_eq!(range.get_value((2, 0)), Some(&Data::Float(42.0)));
    assert_eq!(range.get_value((3, 0)), Some(&Data::Float(5.25)));
    assert_eq!(range.get_value((4, 0)), Some(&Data::Float(12345.0)));
    assert_eq!(
        range.get_value((5, 0)),
        Some(&Data::DateTimeIso("2026-07-17".to_owned()))
    );
    assert_eq!(
        range.get_value((6, 0)),
        Some(&Data::DateTimeIso("2026-07-17T13:14:15".to_owned()))
    );
    assert_eq!(
        range.get_value((7, 0)),
        Some(&Data::Error(calamine::CellErrorType::NA))
    );
    assert_eq!(
        range.get_value((10, 0)),
        Some(&Data::String("value=42".to_owned()))
    );

    let entries = load_entries(&output)?;
    let sheet = entries
        .iter()
        .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .expect("typed worksheet");
    assert!(sheet.contains("<f>SUM(20,22)</f>"));
    assert!(sheet.contains("r=\"A10\"></c>"));
    Ok(())
}

#[test]
fn fills_java_official_simple_template_with_typed_number() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("java-simple.xlsx");
    let output = directory.path().join("java-simple-filled.xlsx");
    write_compressed_java_fixture(
        &template,
        include_str!("fixtures/java-demo-simple.xlsx.gz.b64"),
    )?;

    let mut source: Xlsx<_> = open_workbook(&template).map_err(test_error)?;
    let source_range = source.worksheet_range("Sheet1").map_err(test_error)?;
    let name_coordinate = find_string_coordinate(&source_range, "{name}").expect("name marker");
    let number_coordinate =
        find_string_coordinate(&source_range, "{number}").expect("number marker");

    fill_xlsx_template(
        &template,
        &output,
        &TemplateData::new()
            .with("name", "张三")
            .with("number", 5.2_f64),
    )?;

    let mut result: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    let range = result.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get_value(name_coordinate),
        Some(&Data::String("张三".to_owned()))
    );
    assert_eq!(range.get_value(number_coordinate), Some(&Data::Float(5.2)));
    Ok(())
}

#[test]
fn java_complex_fill_with_table_appends_summary_after_repeated_fill() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("java-complex-table.xlsx");
    let output = directory.path().join("java-complex-table-filled.xlsx");
    write_compressed_java_fixture(
        &template,
        include_str!("fixtures/java-demo-complex-fill-with-table.xlsx.gz.b64"),
    )?;

    let entries = load_entries(&template)?;
    let shared_strings = entries
        .iter()
        .find(|entry| entry.name == "xl/sharedStrings.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .map(shared_string_values)
        .expect("official shared strings");
    let sheet = entries
        .iter()
        .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .expect("official worksheet");
    let marker = FillWrapper::new([TemplateData::new().with("name", "probe")]);
    let (_, _, marker_row, _, _, _) =
        find_collection_row(sheet, &marker, &shared_strings).expect("list marker row");
    let first_data_row = row_index(marker_row).expect("marker row index") - 1;

    let first = [
        TemplateData::new().with("name", "A").with("number", 1),
        TemplateData::new().with("name", "B").with("number", 2),
        TemplateData::new().with("name", "C").with("number", 3),
    ];
    let second = [
        TemplateData::new().with("name", "D").with("number", 4),
        TemplateData::new().with("name", "E").with("number", 5),
        TemplateData::new().with("name", "F").with("number", 6),
    ];
    let mut writer = ExcelTemplateWriter::new(&template, &output)?;
    writer
        .fill_list(&FillWrapper::new(first), FillConfig::new())?
        .fill_list(&FillWrapper::new(second), FillConfig::new())?
        .fill(&TemplateData::new().with("date", "2019年10月9日13:28:28"))?
        .write_rows([vec![
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
            CellValue::String("统计:1000".to_owned()),
        ]])?
        .finish()?;

    let mut workbook: Xlsx<_> = open_workbook(&output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    for (offset, (name, number)) in [
        ("A", 1.0),
        ("B", 2.0),
        ("C", 3.0),
        ("D", 4.0),
        ("E", 5.0),
        ("F", 6.0),
    ]
    .into_iter()
    .enumerate()
    {
        let row = u32::try_from(first_data_row + offset).expect("small row");
        assert_eq!(
            range.get_value((row, 0)),
            Some(&Data::String(name.to_owned()))
        );
        assert_eq!(range.get_value((row, 1)), Some(&Data::Float(number)));
    }
    let summary_row = u32::try_from(first_data_row + 6).expect("small row");
    assert_eq!(
        range.get_value((summary_row, 3)),
        Some(&Data::String("统计:1000".to_owned()))
    );
    Ok(())
}

#[test]
fn fills_shared_string_placeholders_and_preserves_unknown_values() -> Result<()> {
    let (directory, template) = template_fixture()?;
    let output = directory.path().join("filled.xlsx");
    let data = TemplateData::new()
        .with("name", "A&B <Admin>")
        .with("count", 3);
    fill_xlsx_template(&template, &output, &data)?;

    let mut workbook: Xlsx<_> = open_workbook(&output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get_value((0, 0)),
        Some(&Data::String("Hello A&B <Admin>".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 0)),
        Some(&Data::String("Count: 3".to_owned()))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&Data::String("Unknown: {unknown}".to_owned()))
    );

    fill_xlsx_template(
        &output,
        &output,
        &TemplateData::new().with("unknown", "done"),
    )?;
    let mut workbook: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    assert_eq!(
        workbook
            .worksheet_range("Sheet1")
            .map_err(test_error)?
            .get_value((2, 0)),
        Some(&Data::String("Unknown: done".to_owned()))
    );
    Ok(())
}

#[test]
fn package_entries_directories_permissions_and_binary_data_round_trip() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("entries.zip");
    let entries = vec![
        TemplateEntry {
            name: "folder/".to_owned(),
            is_dir: true,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: Vec::new(),
        },
        TemplateEntry {
            name: "folder/data.bin".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Deflated,
            unix_mode: Some(0o644),
            bytes: vec![0, 1, 2, 3],
        },
    ];
    write_entries(&path, &entries)?;
    let actual = load_entries(&path)?;
    assert_eq!(actual.len(), 2);
    assert!(actual[0].is_dir);
    assert_eq!(actual[1].bytes, vec![0, 1, 2, 3]);
    Ok(())
}

#[test]
fn invalid_archives_xml_and_output_paths_return_typed_errors() -> Result<()> {
    let directory = tempdir()?;
    let corrupt = directory.path().join("corrupt.xlsx");
    fs::write(&corrupt, b"not a zip")?;
    assert!(
        fill_xlsx_template(
            &corrupt,
            &directory.path().join("out.xlsx"),
            &TemplateData::new()
        )
        .is_err()
    );

    let invalid_xml = directory.path().join("invalid-xml.xlsx");
    write_entries(
        &invalid_xml,
        &[TemplateEntry {
            name: "bad.xml".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: vec![0xff],
        }],
    )?;
    assert!(
        fill_xlsx_template(
            &invalid_xml,
            &directory.path().join("invalid-out.xlsx"),
            &TemplateData::new()
        )
        .is_err()
    );

    let (_template_directory, template) = template_fixture()?;
    assert!(fill_xlsx_template(&template, directory.path(), &TemplateData::new()).is_err());
    assert!(load_entries(&directory.path().join("missing.xlsx")).is_err());
    assert_eq!(
        format_error("broken").to_string(),
        "excel format error: broken"
    );
    Ok(())
}

#[test]
fn injected_archive_io_failures_cover_all_propagation_boundaries() -> Result<()> {
    let entries = vec![
        TemplateEntry {
            name: "folder/".to_owned(),
            is_dir: true,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: Vec::new(),
        },
        TemplateEntry {
            name: "folder/data.xml".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Stored,
            unix_mode: Some(0o644),
            bytes: b"<value>{name}</value>".to_vec(),
        },
    ];
    let directory = tempdir()?;
    let archive_path = directory.path().join("faults.zip");
    write_entries(&archive_path, &entries)?;
    let bytes = fs::read(&archive_path)?;

    let read_errors = (0..64)
        .filter(|fail_at| {
            load_entries_from(Box::new(FaultyIo::reading(bytes.clone(), *fail_at))).is_err()
        })
        .count();
    let seek_errors = (0..64)
        .filter(|fail_at| {
            load_entries_from(Box::new(FaultyIo::seeking(bytes.clone(), *fail_at))).is_err()
        })
        .count();
    let write_errors = (0..128)
        .filter(|fail_at| {
            write_entries_to(Box::new(FaultyIo::writing(*fail_at)), &entries).is_err()
        })
        .count();
    assert!(read_errors > 1);
    assert!(seek_errors > 1);
    assert!(write_errors > 3);

    let read_only_path = directory.path().join("read-only.bin");
    fs::write(&read_only_path, b"existing")?;
    assert!(write_file_entries(File::open(read_only_path)?, &entries).is_err());

    let mut missing_writer: Option<ArchiveWriter> = None;
    assert!(finish_zip_writer(&mut missing_writer).is_err());
    let mut success = successful_zip_operation;
    assert!(zip_writer_operation(&mut missing_writer, &mut success).is_err());

    let mut active_writer = Some(ZipWriter::new(
        Box::new(Cursor::new(Vec::new())) as Box<dyn WriteSeek>
    ));
    zip_writer_operation(&mut active_writer, &mut success)?;
    let _ = finish_zip_writer(&mut active_writer)?;

    let mut panicking_writer = Some(ZipWriter::new(
        Box::new(Cursor::new(Vec::new())) as Box<dyn WriteSeek>
    ));
    let mut panic_operation =
        |_: &mut ArchiveWriter| -> Result<()> { panic!("injected ZIP panic") };
    assert!(zip_writer_operation(&mut panicking_writer, &mut panic_operation).is_err());
    assert!(panicking_writer.is_none());
    Ok(())
}

#[test]
fn fill_config_and_wrapper_match_java_defaults_and_builders() {
    let rows = vec![TemplateData::new().with("name", "Alice")];
    let unnamed = FillWrapper::new(rows.clone());
    assert_eq!(unnamed.name(), None);
    assert_eq!(unnamed.rows(), rows);

    let named = FillWrapper::named("users", rows.clone());
    assert_eq!(named.name(), Some("users"));
    assert_eq!(named.rows(), rows);
    assert_eq!(FillWrapper::default().rows(), &[]);

    let defaults = FillConfig::default();
    assert_eq!(defaults, FillConfig::new());
    assert_eq!(defaults.get_direction(), FillDirection::Vertical);
    assert!(!defaults.get_force_new_row());
    assert!(defaults.get_auto_style());

    let configured = FillConfig::new()
        .direction(FillDirection::Horizontal)
        .force_new_row(true)
        .auto_style(false);
    assert_eq!(configured.get_direction(), FillDirection::Horizontal);
    assert!(configured.get_force_new_row());
    assert!(!configured.get_auto_style());
}

#[test]
#[allow(clippy::too_many_lines)]
fn stateful_template_writer_matches_java_repeated_and_composite_fill() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("composite-template.xlsx");
    let output = directory.path().join("composite-output.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string(0, 0, "Report {date}")
        .map_err(test_error)?;
    worksheet
        .write_string(0, 1, r"\{date\}")
        .map_err(test_error)?;
    worksheet
        .write_string(1, 0, "{data1.name}")
        .map_err(test_error)?;
    worksheet
        .write_string(3, 0, "{data2.name}")
        .map_err(test_error)?;
    worksheet.write_string(6, 0, "Footer").map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    let horizontal = FillConfig::new().direction(FillDirection::Horizontal);
    let vertical = FillConfig::new().force_new_row(true);
    let mut writer = ExcelTemplateWriter::new(&template, &output)?;
    assert!(!writer.is_finished());
    writer
        .fill(&TemplateData::new().with("date", "old"))?
        .fill_list(
            &FillWrapper::named(
                "data1",
                [
                    TemplateData::new().with("name", "A"),
                    TemplateData::new().with("name", "B"),
                ],
            ),
            horizontal,
        )?
        .fill_list(
            &FillWrapper::named("data1", [TemplateData::new().with("name", "C")]),
            horizontal,
        )?
        .fill_list(
            &FillWrapper::named("data2", [TemplateData::new().with("name", "X")]),
            vertical,
        )?
        .fill_list(
            &FillWrapper::named(
                "data2",
                [
                    TemplateData::new().with("name", "Y"),
                    TemplateData::new().with("name", "Z"),
                ],
            ),
            vertical,
        )?
        .fill_list(&FillWrapper::default(), FillConfig::new())?
        .fill(&TemplateData::new().with("date", 2026))?
        .write_rows([vec![
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
            CellValue::String("统计:1000".to_owned()),
        ]])?;
    assert!(
        writer
            .fill_list(
                &FillWrapper::named("data1", [TemplateData::new().with("name", "invalid")]),
                FillConfig::new(),
            )
            .is_err()
    );
    writer.finish()?;
    writer.finish()?;
    assert!(writer.is_finished());
    assert!(writer.fill(&TemplateData::new()).is_err());
    assert!(writer.write_rows([Vec::<CellValue>::new()]).is_err());
    assert!(
        writer
            .fill_list(&FillWrapper::default(), FillConfig::new())
            .is_err()
    );

    let mut workbook: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get_value((0, 0)),
        Some(&Data::String("Report 2026".to_owned()))
    );
    assert_eq!(
        range.get_value((0, 1)),
        Some(&Data::String("{date}".to_owned()))
    );
    for (column, expected) in [(0_u32, "A"), (1, "B"), (2, "C")] {
        assert_eq!(
            range.get_value((1, column)),
            Some(&Data::String(expected.to_owned()))
        );
    }
    for (row, expected) in [(3_u32, "X"), (4, "Y"), (5, "Z")] {
        assert_eq!(
            range.get_value((row, 0)),
            Some(&Data::String(expected.to_owned()))
        );
    }
    assert_eq!(
        range.get_value((8, 0)),
        Some(&Data::String("Footer".to_owned()))
    );
    assert_eq!(
        range.get_value((9, 3)),
        Some(&Data::String("统计:1000".to_owned()))
    );
    Ok(())
}

#[test]
fn fills_java_official_composite_template_across_all_analysis_cells() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("java-composite.xlsx");
    let output = directory.path().join("java-composite-filled.xlsx");
    write_java_composite_fixture(&template)?;

    let entries = load_entries(&template)?;
    let shared_strings = entries
        .iter()
        .find(|entry| entry.name == "xl/sharedStrings.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .map(shared_string_values)
        .expect("official shared strings");
    let sheet = entries
        .iter()
        .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .expect("official worksheet");
    let data2 = FillWrapper::named(
        "data2",
        [TemplateData::new().with("name", "X").with("number", 10)],
    );
    let (_, _, data2_row, _, _, _) =
        find_collection_row(sheet, &data2, &shared_strings).expect("data2 marker row");
    let filled_data2_row = fill_row_cells(
        data2_row,
        &data2.rows()[0],
        data2.name(),
        &shared_strings,
        true,
    );
    assert!(filled_data2_row.contains("r=\"A9\""));
    assert!(filled_data2_row.contains("r=\"B9\""));

    let horizontal = FillConfig::new().direction(FillDirection::Horizontal);
    let mut writer = ExcelTemplateWriter::new(&template, &output)?;
    for row in [
        TemplateData::new().with("name", "A").with("number", 1),
        TemplateData::new().with("name", "B").with("number", 2),
    ] {
        writer.fill_list(&FillWrapper::named("data1", [row]), horizontal)?;
    }
    for row in [
        TemplateData::new().with("name", "X").with("number", 10),
        TemplateData::new().with("name", "Y").with("number", 20),
    ] {
        writer.fill_list(&FillWrapper::named("data2", [row]), FillConfig::new())?;
    }
    for row in [
        TemplateData::new().with("name", "P").with("number", 100),
        TemplateData::new().with("name", "Q").with("number", 200),
    ] {
        writer.fill_list(&FillWrapper::named("data3", [row]), FillConfig::new())?;
    }
    writer
        .fill(&TemplateData::new().with("date", "2026-07-17"))?
        .finish()?;

    let mut workbook: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    for (coordinate, expected) in [
        ((0, 2), "A"),
        ((0, 3), "B"),
        ((2, 2), "A"),
        ((2, 3), "B"),
        ((4, 0), "时间：2026-07-17"),
        ((8, 0), "X"),
        ((9, 0), "Y"),
        ((10, 3), "P"),
        ((11, 3), "Q"),
    ] {
        assert_eq!(
            range.get_value(coordinate),
            Some(&Data::String(expected.to_owned())),
            "coordinate {coordinate:?}"
        );
    }
    for (coordinate, expected) in [
        ((1, 2), 1.0),
        ((1, 3), 2.0),
        ((3, 2), 1.0),
        ((3, 3), 2.0),
        ((8, 1), 10.0),
        ((9, 1), 20.0),
        ((10, 4), 100.0),
        ((11, 4), 200.0),
    ] {
        assert_eq!(
            range.get_value(coordinate),
            Some(&Data::Float(expected)),
            "coordinate {coordinate:?}"
        );
    }
    Ok(())
}

#[test]
fn expands_vertical_named_rows_and_shifts_footer() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("vertical-template.xlsx");
    let output = directory.path().join("vertical-output.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Name").map_err(test_error)?;
    worksheet
        .write_string(1, 0, "{users.name}")
        .map_err(test_error)?;
    worksheet
        .write_string(1, 1, "Age {users.age}")
        .map_err(test_error)?;
    worksheet
        .write_string(1, 2, "Template static")
        .map_err(test_error)?;
    worksheet.write_string(2, 0, "Footer").map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    let wrapper = FillWrapper::named(
        "users",
        [
            TemplateData::new().with("name", "Alice").with("age", 20),
            TemplateData::new().with("name", "Bob").with("age", 30),
            TemplateData::new().with("name", "Carol").with("age", 40),
        ],
    );
    fill_xlsx_template_list(
        &template,
        &output,
        &wrapper,
        FillConfig::new().force_new_row(true),
    )?;

    let mut workbook: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get_value((1, 0)),
        Some(&Data::String("Alice".to_owned()))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&Data::String("Bob".to_owned()))
    );
    assert_eq!(
        range.get_value((3, 1)),
        Some(&Data::String("Age 40".to_owned()))
    );
    assert_eq!(range.get_value((2, 2)), Some(&Data::Empty));
    assert_eq!(range.get_value((3, 2)), Some(&Data::Empty));
    assert_eq!(
        range.get_value((4, 0)),
        Some(&Data::String("Footer".to_owned()))
    );
    Ok(())
}

#[test]
fn default_vertical_fill_reuses_existing_rows_without_copying_static_cells() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("reuse-template.xlsx");
    let output = directory.path().join("reuse-output.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Name").map_err(test_error)?;
    worksheet
        .write_string(1, 0, "{.name}")
        .map_err(test_error)?;
    worksheet
        .write_string(1, 1, "Template static")
        .map_err(test_error)?;
    worksheet.write_string(2, 0, "old").map_err(test_error)?;
    worksheet
        .write_string(2, 1, "Preserve")
        .map_err(test_error)?;
    worksheet.write_string(3, 0, "Footer").map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    fill_xlsx_template_list(
        &template,
        &output,
        &FillWrapper::new([
            TemplateData::new().with("name", "Alice"),
            TemplateData::new().with("name", "Bob"),
        ]),
        FillConfig::new(),
    )?;

    let mut workbook: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get_value((1, 0)),
        Some(&Data::String("Alice".to_owned()))
    );
    assert_eq!(
        range.get_value((1, 1)),
        Some(&Data::String("Template static".to_owned()))
    );
    assert_eq!(
        range.get_value((2, 0)),
        Some(&Data::String("Bob".to_owned()))
    );
    assert_eq!(
        range.get_value((2, 1)),
        Some(&Data::String("Preserve".to_owned()))
    );
    assert_eq!(
        range.get_value((3, 0)),
        Some(&Data::String("Footer".to_owned()))
    );
    Ok(())
}

#[test]
fn expands_horizontal_unnamed_cells_and_can_drop_style() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("horizontal-template.xlsx");
    let output = directory.path().join("horizontal-output.xlsx");
    let mut workbook = Workbook::new();
    workbook
        .add_worksheet()
        .write_string(0, 0, "{.name}")
        .map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    let wrapper = FillWrapper::new([
        TemplateData::new().with("name", "A"),
        TemplateData::new().with("name", "B"),
        TemplateData::new().with("name", "C"),
    ]);
    fill_xlsx_template_list(
        &template,
        &output,
        &wrapper,
        FillConfig::new()
            .direction(FillDirection::Horizontal)
            .auto_style(false),
    )?;

    let mut workbook: Xlsx<_> = open_workbook(output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(range.get_value((0, 0)), Some(&Data::String("A".to_owned())));
    assert_eq!(range.get_value((0, 1)), Some(&Data::String("B".to_owned())));
    assert_eq!(range.get_value((0, 2)), Some(&Data::String("C".to_owned())));
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn collection_parser_defensive_paths_are_deterministic() {
    let empty = FillWrapper::default();
    let mut no_entries = Vec::new();
    replace_collection_placeholders(&mut no_entries, &empty, FillConfig::new());

    let wrapper = FillWrapper::new([TemplateData::new().with("name", "X")]);
    let mut entries = vec![
        TemplateEntry {
            name: "xl/sharedStrings.xml".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: vec![0xff],
        },
        TemplateEntry {
            name: "xl/worksheets/sheet1.xml".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: vec![0xff],
        },
        TemplateEntry {
            name: "other.xml".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: b"<row/>".to_vec(),
        },
        TemplateEntry {
            name: "xl/worksheets/sheet2.xml".to_owned(),
            is_dir: false,
            compression: CompressionMethod::Stored,
            unix_mode: None,
            bytes: b"<worksheet/>".to_vec(),
        },
    ];
    replace_collection_placeholders(&mut entries, &wrapper, FillConfig::new());

    assert_eq!(shared_string_values("<si"), Vec::<String>::new());
    assert_eq!(shared_string_values("<si>missing"), Vec::<String>::new());
    assert_eq!(text_node_values("<t"), "");
    assert_eq!(text_node_values("<t>missing"), "");
    assert!(expand_vertical_rows("<sheet/>", &wrapper, FillConfig::new(), &[]).is_none());
    assert!(expand_horizontal_cells("<row>", &wrapper, &[]).is_none());
    assert_eq!(
        expand_vertical_rows(
            "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{.name}</t></is></c></row>",
            &wrapper,
            FillConfig::new(),
            &[]
        ),
        Some("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>X</t></is></c></row>".to_owned())
    );
    assert!(find_collection_row("<row>", &wrapper, &[]).is_none());
    assert!(find_collection_cell("<c>", &wrapper, &[]).is_none());
    assert_eq!(
        fill_row_cells("before<c", wrapper.rows().first().unwrap(), None, &[], true),
        "before<c"
    );

    let data = wrapper.rows().first().unwrap();
    assert_eq!(exact_collection_value("{.name", data, None), None);
    assert_eq!(
        exact_collection_value("{other.name}", data, Some("users")),
        None
    );
    assert_eq!(
        exact_collection_value("{usersname}", data, Some("users")),
        None
    );
    assert_eq!(exact_collection_value("{name}", data, None), None);
    assert_eq!(fill_cell("<c", data, None, &[], true), "<c");
    assert_eq!(fill_cell("<c></c>", data, None, &[], true), "<c></c>");
    assert_eq!(
        fill_cell(
            "<c t=\"inlineStr\"><is><t>plain</t></is></c>",
            data,
            None,
            &[],
            true
        ),
        "<c t=\"inlineStr\"><is><t>plain</t></is></c>"
    );
    assert_eq!(
        fill_cell(
            "<c r=\"A1\" s=\"2\" t=\"inlineStr\"><is><t>{.name}</t></is></c>",
            data,
            None,
            &[],
            false
        ),
        "<c r=\"A1\" t=\"inlineStr\"><is><t>X</t></is></c>"
    );
    assert_eq!(
        fill_cell(
            "<c r=\"A1\" s=\"2\" t=\"inlineStr\"><is><t>{.name}</t></is></c>",
            data,
            None,
            &[],
            true
        ),
        "<c r=\"A1\" s=\"2\" t=\"inlineStr\"><is><t>X</t></is></c>"
    );
    assert_eq!(
        fill_cell(
            "<c r=\"A1\" s=\"2\" t=\"inlineStr\"><is><t>Name {.name}</t></is></c>",
            data,
            None,
            &[],
            false
        ),
        "<c r=\"A1\" t=\"inlineStr\"><is><t>Name X</t></is></c>"
    );
    assert_eq!(render_typed_cell("<c", &CellValue::Int(1), true), "<c");
    assert_eq!(
        render_typed_cell(
            "<c r=\"A1\"></c>",
            &CellValue::RichText(easyexcel_core::RichTextStringData::new("rich")),
            true
        ),
        "<c r=\"A1\" t=\"inlineStr\"><is><t>rich</t></is></c>"
    );
    assert_eq!(
        render_typed_cell(
            "<c r=\"A1\"></c>",
            &CellValue::Comment {
                value: Box::new(CellValue::String("commented".to_owned())),
                text: "note".to_owned(),
            },
            true
        ),
        "<c r=\"A1\" t=\"inlineStr\"><is><t>commented</t></is></c>"
    );
    assert_eq!(
        render_typed_cell(
            "<c r=\"A1\"></c>",
            &CellValue::Images {
                value: Box::new(CellValue::Bool(false)),
                images: Vec::new(),
            },
            true
        ),
        "<c r=\"A1\" t=\"b\"><v>0</v></c>"
    );
    assert_eq!(
        render_typed_cell(
            "<c r=\"A1\"></c>",
            &CellValue::Hyperlink {
                url: "https://example.com".to_owned(),
                text: "link".to_owned(),
            },
            true
        ),
        "<c r=\"A1\" t=\"inlineStr\"><is><t>link</t></is></c>"
    );
    assert_eq!(
        render_typed_cell("<c r=\"A1\"></c>", &CellValue::Image(vec![1]), true),
        "<c r=\"A1\"></c>"
    );
    assert_eq!(
        replace_scalar_cells_in_xml("<worksheet><sheetData><c", &TemplateData::new(), &[]),
        "<worksheet><sheetData><c"
    );
    assert_eq!(cell_value("<c t=\"s\"></c>", &[]), None);
    assert_eq!(cell_value("<c t=\"s\"><v>x</v></c>", &[]), None);
    assert_eq!(cell_value("<c t=\"s\"><v>9</v></c>", &[]), None);
    assert_eq!(cell_value("<c></c>", &[]), None);
    assert!(!contains_collection_marker("{other.name}", &wrapper));
}

#[test]
fn collection_coordinate_helpers_and_missing_input_are_deterministic() -> Result<()> {
    let directory = tempdir()?;
    assert_eq!(element_value("", "v"), None);
    assert_eq!(element_value("<v>1", "v"), None);
    assert_eq!(attribute_value("", "r"), None);
    assert_eq!(attribute_value(" r=\"broken", "r"), None);
    assert_eq!(replace_attribute("<c>", "r", "A1"), "<c>");
    assert_eq!(remove_attribute("<c>", "s"), "<c>");
    assert_eq!(shift_rows("tail", 2), "tail");
    assert_eq!(shift_rows("tail", 0), "tail");
    assert_eq!(shift_rows("<row r=\"1\">broken", 2), "<row r=\"1\">broken");
    assert_eq!(shift_row("<c r=\"A1\"/>", 1, 1), "<c r=\"B2\"/>");
    assert_eq!(
        shift_row("<row r=\"x\"></row>", 1, 0),
        "<row r=\"x\"></row>"
    );
    assert!(cell_references(" r=\"broken").is_empty());
    assert!(cell_references(" r=\"abc\"").is_empty());
    assert_eq!(shift_cell_reference("ABC", 1, 0), "ABC");
    assert_eq!(shift_cell_reference("A-invalid", 1, 0), "A-invalid");
    assert_eq!(shift_cell_reference("1", 1, 0), "1");
    assert_eq!(shift_cell_reference("A1x", 1, 0), "A1x");
    assert_eq!(column_name(0), "");
    assert_eq!(column_name(27), "AA");
    assert_eq!(worksheet_max_row("<row"), 0);
    assert_eq!(worksheet_max_row("<row><c/></row>"), 0);
    assert!(append_rows_to_xml("<worksheet/>", &[vec![]]).is_err());
    let mut no_sheet = Vec::new();
    assert!(append_rows_to_first_sheet(&mut no_sheet, &[vec![]]).is_err());
    let mut invalid_sheet = vec![TemplateEntry {
        name: "xl/worksheets/sheet1.xml".to_owned(),
        is_dir: false,
        compression: CompressionMethod::Stored,
        unix_mode: None,
        bytes: vec![0xff],
    }];
    assert!(append_rows_to_first_sheet(&mut invalid_sheet, &[vec![]]).is_err());
    invalid_sheet[0].bytes = vec![0xff];
    assert!(replace_scalar_cells(&mut invalid_sheet, &TemplateData::new()).is_err());
    invalid_sheet[0].bytes = vec![0xff];
    let mut no_sheet_data = vec![TemplateEntry {
        name: "xl/worksheets/sheet1.xml".to_owned(),
        is_dir: false,
        compression: CompressionMethod::Stored,
        unix_mode: None,
        bytes: b"<worksheet/>".to_vec(),
    }];
    assert!(append_rows_to_first_sheet(&mut no_sheet_data, &[vec![]]).is_err());
    no_sheet_data[0].bytes = b"<worksheet/>".to_vec();

    let mut invalid_scalar_writer = ExcelTemplateWriter {
        output: directory.path().join("invalid-scalar.xlsx"),
        entries: invalid_sheet,
        scalar: TemplateData::new(),
        collections: Vec::new(),
        appended_rows: Vec::new(),
        finished: false,
    };
    assert!(invalid_scalar_writer.finish().is_err());
    assert!(!invalid_scalar_writer.is_finished());
    let mut invalid_append_writer = ExcelTemplateWriter {
        output: directory.path().join("invalid-append.xlsx"),
        entries: no_sheet_data,
        scalar: TemplateData::new(),
        collections: Vec::new(),
        appended_rows: vec![vec![]],
        finished: false,
    };
    assert!(invalid_append_writer.finish().is_err());
    assert!(!invalid_append_writer.is_finished());

    let wrapper = FillWrapper::new([TemplateData::new().with("name", "X")]);
    assert!(
        fill_xlsx_template_list(
            &directory.path().join("missing.xlsx"),
            &directory.path().join("out.xlsx"),
            &wrapper,
            FillConfig::new()
        )
        .is_err()
    );
    let (_template_directory, template) = template_fixture()?;
    fill_xlsx_template_list(
        &template,
        &directory.path().join("empty-list-output.xlsx"),
        &FillWrapper::default(),
        FillConfig::new(),
    )?;
    Ok(())
}

#[test]
fn collection_row_merge_helpers_cover_missing_existing_and_inserted_rows() {
    let data = TemplateData::new().with("name", "X");
    let wrapper = FillWrapper::new([data.clone()]);
    let marker = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{.name}</t></is></c></row>";
    assert!(
        expand_vertical_rows(marker, &FillWrapper::default(), FillConfig::new(), &[]).is_none()
    );
    assert!(
        expand_vertical_rows(
            "<row><c r=\"A1\" t=\"inlineStr\"><is><t>{.name}</t></is></c></row>",
            &wrapper,
            FillConfig::new(),
            &[]
        )
        .is_none()
    );
    assert!(
        expand_vertical_rows(
            "<row><c r=\"A1\" t=\"inlineStr\"><is><t>{.name}</t></is></c></row>",
            &wrapper,
            FillConfig::new().force_new_row(true),
            &[]
        )
        .is_none()
    );
    assert_eq!(
        collection_only_row("<row", &data, &wrapper, &[], true, 1),
        "<row"
    );
    assert!(collection_cells("<c", &wrapper, &[]).is_empty());

    let inserted = "<row r=\"2\"><c r=\"A2\"></c></row>";
    assert_eq!(
        upsert_collection_row(
            "<row r=\"3\"><c r=\"A3\"></c></row></sheetData>",
            inserted,
            2
        ),
        format!("{inserted}<row r=\"3\"><c r=\"A3\"></c></row></sheetData>")
    );
    assert_eq!(
        upsert_collection_row(
            "<row r=\"1\"><c r=\"A1\"></c></row></sheetData>",
            inserted,
            2
        ),
        format!("<row r=\"1\"><c r=\"A1\"></c></row>{inserted}</sheetData>")
    );
    assert_eq!(
        upsert_collection_row("<row r=\"1\">broken", inserted, 2),
        format!("<row r=\"1\">broken{inserted}")
    );
    assert_eq!(
        upsert_collection_row("<row r=\"bad\"></row>", inserted, 2),
        format!("<row r=\"bad\"></row>{inserted}")
    );

    assert_eq!(
        merge_collection_cells("<row r=\"2\"></row>", "<row><c><v>1</v></c></row>"),
        "<row r=\"2\"></row>"
    );
    assert_eq!(
        merge_collection_cells(
            "<row r=\"2\"><c r=\"A2\"></c></row>",
            "<row r=\"2\"><c r=\"B2\"></c></row>"
        ),
        "<row r=\"2\"><c r=\"A2\"></c><c r=\"B2\"></c></row>"
    );
    assert_eq!(
        merge_collection_cells(
            "<row r=\"2\"><c r=\"A2\"></c>",
            "<row r=\"2\"><c r=\"B2\"></c></row>"
        ),
        "<row r=\"2\"><c r=\"A2\"></c>"
    );
    assert!(all_cells("<c").is_empty());
    assert_eq!(row_index("<row></row>"), None);
    assert_eq!(row_index("<row r=\"bad\"></row>"), None);
}

#[test]
fn worksheet_metadata_shifts_ranges_formulas_and_recomputes_dimension() {
    let xml = concat!(
        "<worksheet><dimension ref=\"A1:D3\"/><sheetData>",
        "<row r=\"1\"><c r=\"A1\"></c></row>",
        "<row r=\"4\"><c r=\"D4\"><f>SUM(A1:A3)+$B$3+Sheet1!A3+LOG10(A3)</f></c></row>",
        "</sheetData><mergeCells><mergeCell ref=\"B3:C3\"/></mergeCells>",
        "<hyperlinks><hyperlink ref=\"A3\"/></hyperlinks>",
        "<autoFilter ref=\"A1:D3\"/>",
        "<dataValidations><dataValidation sqref=\"A3 B1:B3\"/></dataValidations>",
        "<conditionalFormatting sqref=\"C3:D3\"></conditionalFormatting></worksheet>"
    );
    let shifted = shift_worksheet_metadata(xml, 3, 1);
    assert!(shifted.contains("mergeCell ref=\"B4:C4\""));
    assert!(shifted.contains("hyperlink ref=\"A4\""));
    assert!(shifted.contains("autoFilter ref=\"A1:D4\""));
    assert!(shifted.contains("dataValidation sqref=\"A4 B1:B4\""));
    assert!(shifted.contains("conditionalFormatting sqref=\"C4:D4\""));
    assert!(shifted.contains("SUM(A1:A4)+$B$4+Sheet1!A4+LOG10(A4)"));

    let dimension = update_worksheet_dimension(&shifted);
    assert!(dimension.contains("dimension ref=\"A1:D4\""));
    assert_eq!(shift_worksheet_metadata(xml, 3, 0), xml);
    assert_eq!(update_worksheet_dimension("<worksheet/>"), "<worksheet/>");
}

#[test]
fn force_new_row_pipeline_shifts_real_formula_merge_and_dimension_metadata() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("metadata-template.xlsx");
    let output = directory.path().join("metadata-output.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_string(0, 0, "Name").map_err(test_error)?;
    worksheet
        .write_string(1, 0, "{.name}")
        .map_err(test_error)?;
    worksheet
        .merge_range(2, 1, 2, 2, "Footer", &Format::new())
        .map_err(test_error)?;
    worksheet.write_formula(2, 3, "=A3").map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    fill_xlsx_template_list(
        &template,
        &output,
        &FillWrapper::new([
            TemplateData::new().with("name", "A"),
            TemplateData::new().with("name", "B"),
            TemplateData::new().with("name", "C"),
        ]),
        FillConfig::new().force_new_row(true),
    )?;

    let entries = load_entries(&output)?;
    let worksheet = entries
        .iter()
        .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
        .ok_or_else(|| ExcelError::Format("worksheet fixture is missing".to_owned()))?;
    let xml = std::str::from_utf8(&worksheet.bytes).map_err(test_error)?;
    assert!(xml.contains("dimension ref=\"A1:D5\""));
    assert!(xml.contains("mergeCell ref=\"B5:C5\""));
    assert!(xml.contains("<c r=\"D5\"><f>A5</f>"));
    Ok(())
}

#[test]
fn a1_reference_and_metadata_parsers_reject_malformed_inputs() {
    assert_eq!(parse_cell_reference("$AA$12"), Some((27, 12)));
    assert_eq!(parse_cell_reference("A"), None);
    assert_eq!(parse_cell_reference("1"), None);
    assert_eq!(parse_cell_reference("A0"), None);
    assert_eq!(parse_cell_reference("A1x"), None);
    assert_eq!(parse_cell_reference("XFE1"), None);
    assert_eq!(parse_cell_reference("ZZZZZZZZZZZZZZZZZZZZ1"), None);
    assert_eq!(shift_a1_reference("A2", 3, 2), "A2");
    assert_eq!(shift_a1_reference("bad", 1, 2), "bad");
    assert_eq!(shift_a1_reference("$A$3", 3, 2), "$A$5");
    assert_eq!(shift_reference_list("A1:A3 C3", 3, 1), "A1:A4 C4");

    assert_eq!(shift_formula_elements("<f", 1, 1), "<f");
    assert_eq!(shift_formula_elements("<f>missing", 1, 1), "<f>missing");
    assert_eq!(shift_formula_references("$", 1, 1), "$");
    assert_eq!(shift_formula_references("A3_name+A3x", 1, 1), "A3_name+A3x");
    assert_eq!(shift_formula_references("Sheet1!A3", 3, 1), "Sheet1!A4");
    assert_eq!(shift_formula_references("LOG10(A3)", 3, 1), "LOG10(A4)");

    assert_eq!(
        shift_tag_references("<mergeCell", "mergeCell", "ref", 1, 1),
        "<mergeCell"
    );
    assert_eq!(
        shift_tag_references("<mergeCell/>", "mergeCell", "ref", 1, 1),
        "<mergeCell/>"
    );
    assert_eq!(
        replace_tag_attribute("<x/>", "dimension", "ref", "A1"),
        "<x/>"
    );
    assert_eq!(
        replace_tag_attribute("<dimension", "dimension", "ref", "A1"),
        "<dimension"
    );
    assert_eq!(
        update_worksheet_dimension(
            "<worksheet><dimension ref=\"A1\"/><c></c><c r=\"bad\"></c></worksheet>"
        ),
        "<worksheet><dimension ref=\"A1\"/><c></c><c r=\"bad\"></c></worksheet>"
    );
}
