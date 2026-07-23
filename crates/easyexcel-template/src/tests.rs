use std::fs;
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

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

#[derive(Debug, Default)]
struct SharedOutputState {
    bytes: Vec<u8>,
    fail_write: bool,
    fail_flush: bool,
    flushes: usize,
}

#[derive(Clone, Debug)]
struct SharedOutput(Arc<Mutex<SharedOutputState>>);

impl SharedOutput {
    fn new(fail_write: bool, fail_flush: bool) -> Self {
        Self(Arc::new(Mutex::new(SharedOutputState {
            fail_write,
            fail_flush,
            ..SharedOutputState::default()
        })))
    }
}

impl Write for SharedOutput {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let mut state = self.0.lock().expect("output state lock");
        if state.fail_write {
            return Err(io::Error::other("injected stream write failure"));
        }
        state.bytes.extend_from_slice(buffer);
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        let mut state = self.0.lock().expect("output state lock");
        state.flushes += 1;
        if state.fail_flush {
            Err(io::Error::other("injected stream flush failure"))
        } else {
            Ok(())
        }
    }
}

struct DropReader {
    inner: Cursor<Vec<u8>>,
    dropped: Arc<AtomicBool>,
}

impl DropReader {
    fn new(bytes: Vec<u8>, dropped: Arc<AtomicBool>) -> Self {
        Self {
            inner: Cursor::new(bytes),
            dropped,
        }
    }
}

impl Read for DropReader {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        self.inner.read(buffer)
    }
}

impl Drop for DropReader {
    fn drop(&mut self) {
        self.dropped.store(true, Ordering::SeqCst);
    }
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

fn multi_sheet_template_fixture() -> Result<(TempDir, std::path::PathBuf)> {
    let directory = tempdir()?;
    let path = directory.path().join("multi-sheet-template.xlsx");
    let mut workbook = Workbook::new();
    let summary = workbook.add_worksheet();
    summary.set_name("摘要").map_err(test_error)?;
    summary.write_string(0, 0, "{title}").map_err(test_error)?;

    let details = workbook.add_worksheet();
    details.set_name("明细").map_err(test_error)?;
    details.write_string(0, 0, "{title}").map_err(test_error)?;
    details
        .write_string(1, 0, "{items.name}")
        .map_err(test_error)?;
    details
        .write_string(1, 1, "{items.value}")
        .map_err(test_error)?;

    let untouched = workbook.add_worksheet();
    untouched.set_name("未处理").map_err(test_error)?;
    untouched
        .write_string(0, 0, "{title}")
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

fn synthetic_entry(name: &str, bytes: impl Into<Vec<u8>>) -> TemplateEntry {
    TemplateEntry {
        name: name.to_owned(),
        is_dir: false,
        compression: CompressionMethod::Stored,
        unix_mode: None,
        bytes: bytes.into(),
    }
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
fn template_reader_and_owned_output_follow_java_default_close_lifecycle() -> Result<()> {
    let (_directory, template) = template_fixture()?;
    let input_dropped = Arc::new(AtomicBool::new(false));
    let input = DropReader::new(fs::read(&template)?, Arc::clone(&input_dropped));
    let state = SharedOutput::new(false, false);
    let stream = ExcelOutputStream::new(state.clone());
    let observer = stream.clone();

    let mut writer = ExcelTemplateWriter::from_reader_to_output_stream(input, stream)?;
    assert!(input_dropped.load(Ordering::SeqCst));
    assert!(format!("{writer:?}").contains("owned stream"));
    assert_eq!(
        worksheet_path(&writer.entries, &TemplateSheet::first())?,
        "xl/worksheets/sheet1.xml"
    );
    writer.fill(&TemplateData::new().with("name", "stream"))?;
    assert_eq!(
        writer.sheets[0].scalar.values().get("name"),
        Some(&CellValue::String("stream".to_owned()))
    );
    let shared_strings = writer
        .entries
        .iter()
        .find(|entry| entry.name == "xl/sharedStrings.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .map_or_else(Vec::new, shared_string_values);
    let worksheet = writer
        .entries
        .iter()
        .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
        .and_then(|entry| std::str::from_utf8(&entry.bytes).ok())
        .expect("worksheet XML");
    assert!(
        replace_scalar_cells_in_xml(worksheet, &writer.sheets[0].scalar, &shared_strings)
            .contains("stream")
    );
    writer.finish()?;

    assert!(observer.is_closed());
    let bytes = state.0.lock().expect("output state lock").bytes.clone();
    let mut workbook = Xlsx::new(Cursor::new(bytes)).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    assert_eq!(
        range.get((0, 0)),
        Some(&Data::String("Hello stream".to_owned()))
    );
    Ok(())
}

#[test]
fn template_path_to_owned_output_can_retain_stream() -> Result<()> {
    let (_directory, template) = template_fixture()?;
    let state = SharedOutput::new(false, false);
    let stream = ExcelOutputStream::new(state.clone());
    let observer = stream.clone();
    let mut writer =
        ExcelTemplateWriter::to_output_stream(&template, stream)?.auto_close_stream(false);

    writer.finish()?;
    assert!(!observer.is_closed());
    assert!(observer.with_inner(|_| ()).is_some());
    observer.close()?;
    assert!(observer.is_closed());
    assert!(!state.0.lock().expect("output state lock").bytes.is_empty());
    Ok(())
}

#[test]
fn template_borrowed_output_remains_usable_for_path_and_reader_inputs() -> Result<()> {
    let (directory, template) = template_fixture()?;
    let reader_output = directory.path().join("reader-output.xlsx");
    let mut path_writer =
        ExcelTemplateWriter::from_reader(Cursor::new(fs::read(&template)?), &reader_output)?;
    assert!(format!("{path_writer:?}").contains("path"));
    path_writer.finish()?;
    Xlsx::new(Cursor::new(fs::read(reader_output)?)).map_err(test_error)?;

    let mut first = Cursor::new(Vec::new());
    {
        let mut writer = ExcelTemplateWriter::to_writer(&template, &mut first)?;
        assert!(format!("{writer:?}").contains("borrowed stream"));
        writer.finish()?;
        writer.finish()?;
    }
    let first_bytes = first.get_ref().clone();
    first.write_all(b"caller-owned")?;
    assert!(first.get_ref().ends_with(b"caller-owned"));
    Xlsx::new(Cursor::new(first_bytes)).map_err(test_error)?;

    let mut second = Cursor::new(Vec::new());
    ExcelTemplateWriter::from_reader_to_writer(Cursor::new(fs::read(&template)?), &mut second)?
        .finish()?;
    Xlsx::new(Cursor::new(second.into_inner())).map_err(test_error)?;
    Ok(())
}

#[test]
fn template_stream_failures_are_propagated_and_owned_stream_is_closed() -> Result<()> {
    let (_directory, template) = template_fixture()?;
    let bytes = fs::read(&template)?;
    assert!(
        ExcelTemplateWriter::from_reader(
            FaultyIo::reading(bytes.clone(), 0),
            template.with_extension("read-error.xlsx")
        )
        .is_err()
    );

    let missing = template.with_extension("missing.xlsx");
    let mut constructor_output = Cursor::new(Vec::new());
    assert!(ExcelTemplateWriter::to_writer(&missing, &mut constructor_output).is_err());
    assert!(
        ExcelTemplateWriter::from_reader_to_writer(
            FaultyIo::reading(bytes.clone(), 0),
            &mut constructor_output
        )
        .is_err()
    );
    assert!(
        ExcelTemplateWriter::to_output_stream(
            &missing,
            ExcelOutputStream::new(SharedOutput::new(false, false))
        )
        .is_err()
    );
    assert!(
        ExcelTemplateWriter::from_reader_to_output_stream(
            FaultyIo::reading(bytes.clone(), 0),
            ExcelOutputStream::new(SharedOutput::new(false, false))
        )
        .is_err()
    );
    ExcelTemplateWriter::from_reader(
        FaultyIo::reading(bytes.clone(), usize::MAX),
        template.with_extension("fault-reader-success.xlsx"),
    )?
    .finish()?;
    assert!(
        ExcelTemplateWriter::from_reader(
            Cursor::new(b"not-a-zip".to_vec()),
            template.with_extension("invalid.xlsx")
        )
        .is_err()
    );

    for (fail_write, fail_flush) in [(true, false), (false, true)] {
        let state = SharedOutput::new(fail_write, fail_flush);
        let stream = ExcelOutputStream::new(state);
        let observer = stream.clone();
        let mut writer =
            ExcelTemplateWriter::from_reader_to_output_stream(Cursor::new(bytes.clone()), stream)?;
        assert!(writer.finish().is_err());
        assert!(observer.is_closed());
    }

    for (fail_write, fail_flush) in [(true, false), (false, true)] {
        let mut output = SharedOutput::new(fail_write, fail_flush);
        let mut writer = ExcelTemplateWriter::to_writer(&template, &mut output)?;
        assert!(writer.finish().is_err());
    }

    let entries = load_entries(&template)?;
    let wrong_type = write_entries_to(Box::new(FaultyIo::writing(usize::MAX)), &entries)?;
    assert!(archive_output_bytes(wrong_type).is_err());

    let invalid_entries = [TemplateEntry {
        name: "invalid.bin".to_owned(),
        is_dir: false,
        compression: CompressionMethod::AES,
        unix_mode: None,
        bytes: vec![1],
    }];
    assert!(encode_entries(&invalid_entries).is_err());
    let mut borrowed = Cursor::new(Vec::new());
    assert!(
        write_entries_to_output(
            &mut TemplateOutput::Borrowed(&mut borrowed),
            &invalid_entries,
            true
        )
        .is_err()
    );
    let invalid_stream = ExcelOutputStream::new(SharedOutput::new(false, false));
    let invalid_observer = invalid_stream.clone();
    assert!(
        write_entries_to_output(
            &mut TemplateOutput::Owned(Box::new(invalid_stream)),
            &invalid_entries,
            true
        )
        .is_err()
    );
    assert!(invalid_observer.is_closed());
    Ok(())
}

#[test]
fn stateful_template_writer_isolates_scalar_list_and_rows_by_sheet() -> Result<()> {
    let (directory, template) = multi_sheet_template_fixture()?;
    let output = directory.path().join("multi-sheet-filled.xlsx");
    let details = TemplateSheet::name("明细");
    let rows = FillWrapper::named(
        "items",
        [
            TemplateData::new().with("name", "A").with("value", 1),
            TemplateData::new().with("name", "B").with("value", 2),
        ],
    );

    let mut writer = ExcelTemplateWriter::new(&template, &output)?;
    writer
        .fill(&TemplateData::new().with("title", "首页"))?
        .fill_on_sheet(&details, &TemplateData::new().with("title", "详情"))?
        .fill_list_on_sheet(&details, &rows, FillConfig::new())?
        .fill_on_sheet(
            &TemplateSheet::index(1),
            &TemplateData::new().with("title", "详情覆盖"),
        )?
        .write_rows_on_sheet(
            &TemplateSheet::index(1),
            [vec![
                CellValue::String("合计".to_owned()),
                CellValue::Int(3),
            ]],
        )?
        .finish()?;

    let mut workbook: Xlsx<_> = open_workbook(&output).map_err(test_error)?;
    let summary = workbook.worksheet_range("摘要").map_err(test_error)?;
    assert_eq!(summary.get((0, 0)), Some(&Data::String("首页".to_owned())));
    let details = workbook.worksheet_range("明细").map_err(test_error)?;
    assert_eq!(
        details.get((0, 0)),
        Some(&Data::String("详情覆盖".to_owned()))
    );
    assert_eq!(details.get((1, 0)), Some(&Data::String("A".to_owned())));
    assert_eq!(details.get((1, 1)), Some(&Data::Float(1.0)));
    assert_eq!(details.get((2, 0)), Some(&Data::String("B".to_owned())));
    assert_eq!(details.get((2, 1)), Some(&Data::Float(2.0)));
    assert_eq!(details.get((3, 0)), Some(&Data::String("合计".to_owned())));
    assert_eq!(details.get((3, 1)), Some(&Data::Float(3.0)));
    let untouched = workbook.worksheet_range("未处理").map_err(test_error)?;
    assert_eq!(
        untouched.get((0, 0)),
        Some(&Data::String("{title}".to_owned()))
    );
    Ok(())
}

#[test]
fn repeated_fill_reuses_java_cursor_when_direction_changes() -> Result<()> {
    let directory = tempdir().map_err(test_error)?;
    let template = directory.path().join("direction-change-template.xlsx");
    let output = directory.path().join("direction-change-filled.xlsx");
    let mut workbook = Workbook::new();
    workbook
        .add_worksheet()
        .set_name("纵转横")
        .map_err(test_error)?
        .write_string(0, 0, "{items.name}")
        .map_err(test_error)?;
    workbook
        .add_worksheet()
        .set_name("横转纵")
        .map_err(test_error)?
        .write_string(0, 0, "{items.name}")
        .map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    let first = FillWrapper::named(
        "items",
        [
            TemplateData::new().with("name", "A"),
            TemplateData::new().with("name", "B"),
        ],
    );
    let second = FillWrapper::named(
        "items",
        [
            TemplateData::new().with("name", "C"),
            TemplateData::new().with("name", "D"),
        ],
    );
    let mut writer = ExcelTemplateWriter::new(&template, &output)?;
    writer
        .fill_list_on_sheet(
            &TemplateSheet::name("纵转横"),
            &first,
            FillConfig::new(),
        )?
        .fill_list_on_sheet(
            &TemplateSheet::name("纵转横"),
            &second,
            FillConfig::new().direction(FillDirection::Horizontal),
        )?
        .fill_list_on_sheet(
            &TemplateSheet::name("横转纵"),
            &first,
            FillConfig::new().direction(FillDirection::Horizontal),
        )?
        .fill_list_on_sheet(
            &TemplateSheet::name("横转纵"),
            &second,
            FillConfig::new(),
        )?
        .finish()?;

    let mut workbook: Xlsx<_> = open_workbook(&output).map_err(test_error)?;
    let vertical_then_horizontal = workbook
        .worksheet_range("纵转横")
        .map_err(test_error)?;
    assert_eq!(
        vertical_then_horizontal.get((0, 0)),
        Some(&Data::String("A".to_owned()))
    );
    assert_eq!(
        vertical_then_horizontal.get((1, 0)),
        Some(&Data::String("B".to_owned()))
    );
    assert_eq!(
        vertical_then_horizontal.get((0, 2)),
        Some(&Data::String("C".to_owned()))
    );
    assert_eq!(
        vertical_then_horizontal.get((0, 3)),
        Some(&Data::String("D".to_owned()))
    );

    let horizontal_then_vertical = workbook
        .worksheet_range("横转纵")
        .map_err(test_error)?;
    assert_eq!(
        horizontal_then_vertical.get((0, 0)),
        Some(&Data::String("A".to_owned()))
    );
    assert_eq!(
        horizontal_then_vertical.get((0, 1)),
        Some(&Data::String("B".to_owned()))
    );
    assert_eq!(
        horizontal_then_vertical.get((2, 0)),
        Some(&Data::String("C".to_owned()))
    );
    assert_eq!(
        horizontal_then_vertical.get((3, 0)),
        Some(&Data::String("D".to_owned()))
    );
    Ok(())
}

#[test]
fn repeated_fill_applies_each_calls_force_row_and_auto_style_config() -> Result<()> {
    let directory = tempdir().map_err(test_error)?;
    let template = directory.path().join("config-change-template.xlsx");
    let output = directory.path().join("config-change-filled.xlsx");
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .write_string_with_format(0, 0, "{items.name}", &Format::new().set_bold())
        .map_err(test_error)?;
    worksheet
        .write_string(1, 0, "Footer")
        .map_err(test_error)?;
    workbook.save(&template).map_err(test_error)?;

    let mut writer = ExcelTemplateWriter::new(&template, &output)?;
    writer
        .fill_list(
            &FillWrapper::named("items", [TemplateData::new().with("name", "A")]),
            FillConfig::new(),
        )?
        .fill_list(
            &FillWrapper::named(
                "items",
                [
                    TemplateData::new().with("name", "B"),
                    TemplateData::new().with("name", "C"),
                ],
            ),
            FillConfig::new()
                .force_new_row(true)
                .auto_style(false),
        )?
        .finish()?;

    let mut workbook: Xlsx<_> = open_workbook(&output).map_err(test_error)?;
    let range = workbook.worksheet_range("Sheet1").map_err(test_error)?;
    for (row, value) in [(0, "A"), (1, "B"), (2, "C"), (3, "Footer")] {
        assert_eq!(
            range.get((row, 0)),
            Some(&Data::String(value.to_owned()))
        );
    }

    let entries = load_entries(&output)?;
    let sheet = entries
        .iter()
        .find(|entry| entry.name == "xl/worksheets/sheet1.xml")
        .expect("sheet1 exists");
    let xml = std::str::from_utf8(&sheet.bytes).map_err(test_error)?;
    let style = |reference| {
        all_cells(xml)
            .into_iter()
            .find(|(_, _, cell)| attribute_value(cell, "r") == Some(reference))
            .and_then(|(_, _, cell)| attribute_value(cell, "s"))
    };
    assert!(style("A1").is_some());
    assert_eq!(style("A2"), None);
    assert_eq!(style("A3"), None);
    Ok(())
}

#[test]
fn collection_cursor_defensive_paths_and_shifted_cached_templates_are_covered() -> Result<()> {
    let wrapper = FillWrapper::named(
        "items",
        [TemplateData::new().with("name", "value")],
    );
    let fill = PendingCollectionFill {
        wrapper: wrapper.clone(),
        config: FillConfig::new(),
        order: 0,
    };
    assert!(replace_collection_fills_in_sheet(&mut [], "missing.xml", &[]).is_ok());
    assert!(matches!(
        replace_collection_fills_in_sheet(&mut [], "missing.xml", std::slice::from_ref(&fill)),
        Err(ExcelError::Format(message)) if message.contains("worksheet part")
    ));
    let mut invalid_utf8 = vec![synthetic_entry("xl/worksheets/sheet1.xml", vec![0xff])];
    assert!(replace_collection_fills_in_sheet(
        &mut invalid_utf8,
        "xl/worksheets/sheet1.xml",
        std::slice::from_ref(&fill)
    )
    .is_err());
    let worksheet_without_marker =
        r#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>plain</t></is></c></row></sheetData></worksheet>"#;
    let mut entries_without_marker = vec![synthetic_entry(
        "xl/worksheets/sheet1.xml",
        worksheet_without_marker.as_bytes().to_vec(),
    )];
    let fill_without_marker = PendingCollectionFill {
        wrapper: wrapper.clone(),
        config: FillConfig::new().force_new_row(true),
        order: 1,
    };
    replace_collection_fills_in_sheet(
        &mut entries_without_marker,
        "xl/worksheets/sheet1.xml",
        std::slice::from_ref(&fill_without_marker),
    )?;
    assert_eq!(
        entries_without_marker[0].bytes,
        worksheet_without_marker.as_bytes()
    );

    assert!(collection_template_cells("<row", &wrapper, &[]).is_empty());
    assert!(
        collection_template_cells(
            r#"<row><c t="inlineStr"><is><t>{items.name}</t></is></c></row>"#,
            &wrapper,
            &[]
        )
        .is_empty()
    );
    assert!(
        collection_template_cells(
            r#"<row><c r="bad" t="inlineStr"><is><t>{items.name}</t></is></c></row>"#,
            &wrapper,
            &[]
        )
        .is_empty()
    );
    assert_eq!(row_tag_with_reference("<row>", 7), "<row>");
    assert!(validate_collection_target(1_048_576, 0).is_err());
    assert!(validate_collection_target(0, 16_384).is_err());
    assert_eq!(last_worksheet_row("<row"), None);
    assert_eq!(
        last_worksheet_row(r#"<row r="2"></row><row r="5"></row>"#),
        Some(4)
    );
    assert_eq!(shift_worksheet_rows_after("<row", 0, 1), "<row");
    assert_eq!(
        shift_worksheet_rows_after(r#"<row r="1"></row>"#, 1, 1),
        r#"<row r="1"></row>"#
    );

    let worksheet = r#"<worksheet><dimension ref="A1:A5"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{a.name}</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>{b.name}</t></is></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>Footer</t></is></c></row></sheetData></worksheet>"#;
    let mut entries = vec![synthetic_entry(
        "xl/worksheets/sheet1.xml",
        worksheet.as_bytes().to_vec(),
    )];
    let fills = [
        PendingCollectionFill {
            wrapper: FillWrapper::named(
                "b",
                [TemplateData::new().with("name", "B1")],
            ),
            config: FillConfig::new(),
            order: 0,
        },
        PendingCollectionFill {
            wrapper: FillWrapper::named(
                "a",
                [
                    TemplateData::new().with("name", "A1"),
                    TemplateData::new().with("name", "A2"),
                ],
            ),
            config: FillConfig::new().force_new_row(true),
            order: 1,
        },
        PendingCollectionFill {
            wrapper: FillWrapper::named(
                "b",
                [TemplateData::new().with("name", "B2")],
            ),
            config: FillConfig::new(),
            order: 2,
        },
    ];
    replace_collection_fills_in_sheet(
        &mut entries,
        "xl/worksheets/sheet1.xml",
        &fills,
    )?;
    let xml = std::str::from_utf8(&entries[0].bytes).map_err(test_error)?;
    assert!(xml.contains("A2"));
    assert!(xml.contains("B2"));
    assert!(xml.contains(r#"ref="A1:A6""#));
    Ok(())
}

#[test]
fn template_sheet_selection_reports_missing_names_and_indexes() -> Result<()> {
    let (directory, template) = multi_sheet_template_fixture()?;
    assert_eq!(TemplateSheet::default(), TemplateSheet::first());
    assert!(same_sheet(
        &TemplateSheet::first(),
        &TemplateSheet::index(0)
    ));
    assert!(!same_sheet(
        &TemplateSheet::name("摘要"),
        &TemplateSheet::index(0)
    ));
    assert!(same_sheet(
        &TemplateSheet::index(2),
        &TemplateSheet::index(2)
    ));
    assert!(!same_sheet(
        &TemplateSheet::index(1),
        &TemplateSheet::index(2)
    ));

    for (sheet, name) in [
        (TemplateSheet::index(99), "missing-index.xlsx"),
        (TemplateSheet::name("不存在"), "missing-name.xlsx"),
    ] {
        let mut writer = ExcelTemplateWriter::new(&template, directory.path().join(name))?;
        writer.fill_on_sheet(&sheet, &TemplateData::new().with("title", "x"))?;
        assert!(matches!(writer.finish(), Err(ExcelError::SheetNotFound(_))));
        assert!(!writer.is_finished());
    }

    let rows = FillWrapper::named(
        "items",
        [TemplateData::new().with("name", "A").with("value", 1)],
    );
    let mut writer = ExcelTemplateWriter::new(
        &template,
        directory.path().join("conflicting-sheet-alias.xlsx"),
    )?;
    writer
        .fill_list_on_sheet(&TemplateSheet::name("明细"), &rows, FillConfig::new())?
        .fill_list_on_sheet(
            &TemplateSheet::index(1),
            &rows,
            FillConfig::new().direction(FillDirection::Horizontal),
        )?;
    writer.finish()?;
    assert!(writer.is_finished());

    let mut writer = ExcelTemplateWriter::new(
        &template,
        directory.path().join("distinct-sheet-alias.xlsx"),
    )?;
    writer
        .fill_list_on_sheet(&TemplateSheet::name("明细"), &rows, FillConfig::new())?
        .fill_list_on_sheet(
            &TemplateSheet::index(1),
            &FillWrapper::named("others", [TemplateData::new().with("name", "B")]),
            FillConfig::new(),
        )?;
    let resolved = writer.resolved_sheet_fills()?;
    assert_eq!(resolved.len(), 2);
    assert_eq!(resolved[1].collections.len(), 2);

    let mut writer =
        ExcelTemplateWriter::new(&template, directory.path().join("merged-sheet-alias.xlsx"))?;
    writer
        .fill_list_on_sheet(&TemplateSheet::name("明细"), &rows, FillConfig::new())?
        .fill_list_on_sheet(&TemplateSheet::index(1), &rows, FillConfig::new())?;
    let resolved = writer.resolved_sheet_fills()?;
    assert_eq!(resolved[1].collections.len(), 2);
    assert_eq!(
        resolved[1]
            .collections
            .iter()
            .map(|fill| fill.wrapper.rows().len())
            .sum::<usize>(),
        2
    );
    Ok(())
}

#[test]
fn worksheet_part_resolution_covers_relationship_and_fallback_failures() {
    let workbook = br#"<workbook><sheets><sheet name="Data" r:id="rId1"/></sheets></workbook>"#;
    let missing_relationship = vec![
        synthetic_entry("xl/workbook.xml", workbook.to_vec()),
        synthetic_entry("xl/_rels/workbook.xml.rels", b"<Relationships/>".to_vec()),
    ];
    assert!(matches!(
        worksheet_path(&missing_relationship, &TemplateSheet::first()),
        Err(ExcelError::Format(message)) if message.contains("relationship rId1")
    ));

    let missing_part = vec![
        synthetic_entry("xl/workbook.xml", workbook.to_vec()),
        synthetic_entry(
            "xl/_rels/workbook.xml.rels",
            br#"<Relationships><Relationship Id="rId1" Target="worksheets/missing.xml"/></Relationships>"#.to_vec(),
        ),
    ];
    assert!(matches!(
        worksheet_path(&missing_part, &TemplateSheet::name("Data")),
        Err(ExcelError::Format(message)) if message.contains("worksheet part")
    ));

    for entries in [
        vec![
            synthetic_entry("xl/workbook.xml", vec![0xff]),
            synthetic_entry("xl/_rels/workbook.xml.rels", b"<Relationships/>".to_vec()),
        ],
        vec![
            synthetic_entry("xl/workbook.xml", workbook.to_vec()),
            synthetic_entry("xl/_rels/workbook.xml.rels", vec![0xff]),
        ],
    ] {
        assert!(matches!(
            worksheet_path(&entries, &TemplateSheet::first()),
            Err(ExcelError::Format(_))
        ));
    }
    let invalid_target = vec![
        synthetic_entry("xl/workbook.xml", workbook.to_vec()),
        synthetic_entry(
            "xl/_rels/workbook.xml.rels",
            br#"<Relationships><Relationship Id="rId1" Target="../../outside.xml"/></Relationships>"#.to_vec(),
        ),
    ];
    assert!(matches!(
        worksheet_path(&invalid_target, &TemplateSheet::first()),
        Err(ExcelError::Format(message)) if message.contains("escapes package root")
    ));
    assert!(
        workbook_sheets(
            r#"<sheets><sheet name="missing-id"/><sheet r:id="missing-name"/></sheets>"#
        )
        .is_empty()
    );

    let fallback = vec![synthetic_entry(
        "xl/worksheets/custom.xml",
        b"<worksheet/>".to_vec(),
    )];
    assert_eq!(
        worksheet_path(&fallback, &TemplateSheet::index(0)).expect("fallback index"),
        "xl/worksheets/custom.xml"
    );
    assert!(matches!(
        worksheet_path(&fallback, &TemplateSheet::name("Data")),
        Err(ExcelError::SheetNotFound(name)) if name == "Data"
    ));
    assert!(matches!(
        worksheet_path(&fallback, &TemplateSheet::index(1)),
        Err(ExcelError::SheetNotFound(index)) if index == "1"
    ));

    assert_eq!(
        normalize_workbook_target("../worksheets/sheet.xml").expect("relative target"),
        "worksheets/sheet.xml"
    );
    assert!(normalize_workbook_target("../../outside.xml").is_err());
    assert!(normalize_workbook_target("/").is_err());
    assert_eq!(
        xml_elements("<sheets><sheet name=\"A\"/><sheet", "sheet").collect::<Vec<_>>(),
        vec!["<sheet name=\"A\"/>"]
    );
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
    // `<cols>` / `<col>` must not be treated as cells (complex.xlsx regression).
    let cols_xml = concat!(
        r#"<worksheet><cols><col min="1" max="1" width="10"/></cols>"#,
        r#"<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>{date}</t></is></c>"#,
        r#"<c r="B1" s="1"/></row></sheetData></worksheet>"#
    );
    let filled_cols = replace_scalar_cells_in_xml(
        cols_xml,
        &TemplateData::new().with("date", "2019年10月9日13:28:28"),
        &[],
    );
    assert!(
        filled_cols.contains("<cols>") && filled_cols.contains("</cols>"),
        "cols section must survive scalar fill: {filled_cols}"
    );
    assert!(
        filled_cols.contains("2019年10月9日13:28:28"),
        "scalar date must be filled: {filled_cols}"
    );
    assert!(
        filled_cols.contains(r#"<c r="B1" s="1"/>"#),
        "self-closing cells must survive: {filled_cols}"
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
        output: TemplateOutput::Path(directory.path().join("invalid-scalar.xlsx")),
        entries: invalid_sheet,
        sheets: vec![PendingSheetFill::new(TemplateSheet::first())],
        next_collection_order: 0,
        finished: false,
        auto_close_stream: true,
    };
    assert!(invalid_scalar_writer.finish().is_err());
    assert!(!invalid_scalar_writer.is_finished());
    let mut invalid_append_writer = ExcelTemplateWriter {
        output: TemplateOutput::Path(directory.path().join("invalid-append.xlsx")),
        entries: no_sheet_data,
        sheets: vec![PendingSheetFill {
            sheet: TemplateSheet::first(),
            scalar: TemplateData::new(),
            collections: Vec::new(),
            appended_rows: vec![vec![]],
        }],
        next_collection_order: 0,
        finished: false,
        auto_close_stream: true,
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
