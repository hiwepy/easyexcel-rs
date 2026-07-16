use std::fs;
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};

use calamine::{Data, Reader, Xlsx, open_workbook};
use rust_xlsxwriter::Workbook;
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

#[test]
fn template_data_and_xml_escaping_are_deterministic() {
    let mut data = TemplateData::new().with("name", "Alice").with("count", 2);
    assert_eq!(data.insert("name", "Bob"), Some("Alice".to_owned()));
    assert_eq!(data.insert("new", "value"), None);
    assert_eq!(data.values().get("name").map(String::as_str), Some("Bob"));
    assert_eq!(escape_xml("<&>\"' text"), "&lt;&amp;&gt;&quot;&apos; text");
    assert_eq!(
        replace_placeholders(
            "{a}-{missing}-{b}",
            &BTreeMap::from([
                ("a".to_owned(), "<".to_owned()),
                ("b".to_owned(), "&".to_owned())
            ])
        ),
        "&lt;-{missing}-&amp;"
    );
    assert_eq!(TemplateData::default(), TemplateData::new());
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
    assert_eq!(
        range.get_value((4, 0)),
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
    assert_eq!(cell_value("<c t=\"s\"></c>", &[]), None);
    assert_eq!(cell_value("<c t=\"s\"><v>x</v></c>", &[]), None);
    assert_eq!(cell_value("<c t=\"s\"><v>9</v></c>", &[]), None);
    assert_eq!(cell_value("<c></c>", &[]), None);
    assert!(!contains_collection_marker("{other.name}", &wrapper));
}

#[test]
fn collection_coordinate_helpers_and_missing_input_are_deterministic() -> Result<()> {
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

    let directory = tempdir()?;
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
    Ok(())
}
