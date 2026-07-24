use std::collections::BTreeMap;
use std::fs;
use std::io::{self, Cursor, Read as _, Write};
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};

use bigdecimal::BigDecimal;
use chrono::NaiveDate;
use tempfile::tempdir;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

use super::*;

/// Reads a ZIP entry from an XLSX package as UTF-8 text (integration asserts).
fn zip_entry_text(path: &Path, name: &str) -> Result<String> {
    let file = fs::File::open(path)?;
    let mut archive =
        ZipArchive::new(file).map_err(|error| ExcelError::Format(error.to_string()))?;
    let mut entry = archive
        .by_name(name)
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    let mut value = String::new();
    entry.read_to_string(&mut value)?;
    Ok(value)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Value(String);

impl ExcelRow for Value {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Value", Some(0), 0, None)];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        Ok(Self(
            row.cell(&Self::schema()[0])
                .map_or_else(String::new, CellValue::as_text),
        ))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![CellValue::String(self.0.clone())])
    }
}

#[test]
fn easy_excel_inherits_factory_entry_points_through_the_same_rust_type() {
    assert_eq!(
        std::any::TypeId::of::<EasyExcel>(),
        std::any::TypeId::of::<EasyExcelFactory>()
    );
    assert_eq!(EasyExcelFactory::writer_table(3).table_no(), 3);
    assert_eq!(
        EasyExcelFactory::writer_sheet_index::<Value>(4)
            .options()
            .sheet_index,
        Some(4)
    );
}

#[test]
fn easy_excel_factory_builds_all_unbound_sheet_and_table_overloads() {
    let default_read_sheet = EasyExcelFactory::read_sheet().build();
    assert!(!default_read_sheet.has_sheet_no());
    assert!(default_read_sheet.sheet_name().is_empty());

    let indexed_read_sheet = EasyExcelFactory::read_sheet_index(2).build();
    assert!(indexed_read_sheet.has_sheet_no());
    assert_eq!(indexed_read_sheet.sheet_no(), 2);

    let named_read_sheet = EasyExcelFactory::read_sheet_name("Named").build();
    assert!(!named_read_sheet.has_sheet_no());
    assert_eq!(named_read_sheet.sheet_name(), "Named");

    let combined_read_sheet = EasyExcelFactory::read_sheet_with(3, "Combined").build();
    assert_eq!(combined_read_sheet.sheet_no(), 3);
    assert_eq!(combined_read_sheet.sheet_name(), "Combined");

    let default_write_sheet = EasyExcelFactory::writer_sheet_builder().build();
    assert_eq!(default_write_sheet.sheet_no(), 0);
    assert!(default_write_sheet.sheet_name().is_empty());

    let indexed_write_sheet = EasyExcelFactory::writer_sheet_builder_index(4).build();
    assert_eq!(indexed_write_sheet.sheet_no(), 4);

    let named_write_sheet = EasyExcelFactory::writer_sheet_builder_name("Output").build();
    assert_eq!(named_write_sheet.sheet_name(), "Output");

    let combined_write_sheet =
        EasyExcelFactory::writer_sheet_builder_with(5, "CombinedOutput").build();
    assert_eq!(combined_write_sheet.sheet_no(), 5);
    assert_eq!(combined_write_sheet.sheet_name(), "CombinedOutput");

    assert_eq!(
        EasyExcelFactory::writer_table_builder_default()
            .build()
            .table_no(),
        0
    );
    assert_eq!(
        EasyExcelFactory::writer_table_builder(6).build().table_no(),
        6
    );
}

#[test]
fn easy_excel_factory_input_stream_uses_the_real_xlsx_reader_and_cleans_up() -> Result<()> {
    let directory = tempdir()?;
    let source = directory.path().join("factory-stream.xlsx");
    EasyExcelFactory::write::<Value>(&source)
        .need_head(false)
        .do_write([Value("from-stream".to_owned())])?;
    let bytes = fs::read(source)?;

    let events = Arc::new(Mutex::new(Vec::new()));
    let listener = OrderedReadListener {
        name: "stream",
        events: Arc::clone(&events),
    };
    let builder =
        EasyExcelFactory::reader_from_input_stream(Cursor::new(bytes))?.head_row_number(0);
    let temporary_path = builder
        .file
        .as_ref()
        .expect("input stream builder must expose its materialised path")
        .to_owned();
    assert!(temporary_path.exists());

    let builder = builder.register_read_listener::<Value, _>(listener);
    let mut reader = builder.build()?;
    assert!(reader.has_temporary_input());
    reader.read_all()?;
    reader.finish();
    assert!(!reader.has_temporary_input());

    assert_eq!(
        *events.lock().expect("factory stream events lock"),
        vec!["stream:from-stream"]
    );
    assert!(
        !temporary_path.exists(),
        "temporary input must be deleted immediately by finish"
    );
    Ok(())
}

#[test]
fn easy_excel_factory_input_stream_is_cleaned_when_analysis_fails() -> Result<()> {
    struct RejectRow;

    impl ReadListener<Value> for RejectRow {
        fn invoke(&mut self, _data: Value, _context: &AnalysisContext) -> Result<()> {
            Err(ExcelError::Format("listener rejected row".to_owned()))
        }
    }

    let directory = tempdir()?;
    let source = directory.path().join("factory-stream-error.xlsx");
    EasyExcelFactory::write::<Value>(&source)
        .need_head(false)
        .do_write([Value("reject-me".to_owned())])?;
    let builder = EasyExcelFactory::reader_from_input_stream(Cursor::new(fs::read(source)?))?
        .head_row_number(0);
    let temporary_path = builder
        .file
        .as_ref()
        .expect("materialised input path")
        .to_owned();
    let mut reader = builder.build(RejectRow)?;

    let error = reader.read_all().expect_err("listener must reject the row");
    assert!(error.to_string().contains("listener rejected row"));
    assert!(!reader.has_temporary_input());
    assert!(
        !temporary_path.exists(),
        "analysis failure must run finish and delete temporary input"
    );
    Ok(())
}

#[test]
fn easy_excel_factory_detects_an_encrypted_xlsx_input_stream() -> Result<()> {
    let directory = tempdir()?;
    let source = directory.path().join("factory-encrypted.xlsx");
    EasyExcelFactory::write::<Value>(&source)
        .password("stream-secret")
        .need_head(false)
        .do_write([Value("encrypted-stream".to_owned())])?;
    let bytes = fs::read(source)?;
    assert!(bytes.starts_with(b"\xD0\xCF\x11\xE0"));

    let events = Arc::new(Mutex::new(Vec::new()));
    let listener = OrderedReadListener {
        name: "encrypted",
        events: Arc::clone(&events),
    };
    let mut reader = EasyExcelFactory::reader_from_input_stream(Cursor::new(bytes))?
        .password("stream-secret")
        .head_row_number(0)
        .build(listener)?;
    reader.read_all()?;
    assert_eq!(
        *events.lock().expect("encrypted stream events lock"),
        vec!["encrypted:encrypted-stream"]
    );
    Ok(())
}

#[test]
fn easy_excel_factory_path_and_output_stream_builders_execute_real_writes() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("factory-path.xlsx");
    EasyExcelFactory::writer_to_path(&path)
        .sheet_name("Path")
        .expect("path-backed writer sheet")
        .need_head(false)
        .do_write([Value("path".to_owned())])?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&path)
            .sheet("Path")
            .head_row_number(0)
            .do_read_sync()?,
        vec![Value("path".to_owned())]
    );

    let output = ExcelOutputStream::new(Cursor::new(Vec::<u8>::new()));
    let inspect = output.clone();
    EasyExcelFactory::writer()
        .auto_close_stream(false)
        .output_stream(output)
        .sheet_name("Stream")
        .need_head(false)
        .do_write([Value("output-stream".to_owned())])?;
    let bytes = inspect
        .with_inner(|cursor| cursor.get_ref().clone())
        .expect("auto_close_stream(false) keeps output inspectable");
    assert!(bytes.starts_with(b"PK"));

    let observed = Arc::new(Mutex::new(Vec::new()));
    let listener = OrderedReadListener {
        name: "output",
        events: Arc::clone(&observed),
    };
    EasyExcelFactory::reader_from_input_stream(Cursor::new(bytes))?
        .head_row_number(0)
        .sheet_name("Stream")
        .build(listener)?
        .read_all()?;
    assert_eq!(
        *observed.lock().expect("factory output events lock"),
        vec!["output:output-stream"]
    );
    Ok(())
}

#[derive(Clone)]
struct FallibleValue {
    value: &'static str,
    fail: bool,
}

#[derive(Default)]
struct FacadeProbeWrite {
    bytes: Vec<u8>,
    fail_write: bool,
    fail_flush: bool,
    fail_flush_at: Option<usize>,
    flushes: usize,
}

impl Write for FacadeProbeWrite {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.fail_write {
            Err(io::Error::other("injected facade write failure"))
        } else {
            self.bytes.extend_from_slice(buffer);
            Ok(buffer.len())
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        let flush = self.flushes;
        self.flushes += 1;
        if self.fail_flush || self.fail_flush_at == Some(flush) {
            Err(io::Error::other("injected facade flush failure"))
        } else {
            Ok(())
        }
    }
}

#[derive(Clone)]
struct ToggleFacadeWrite {
    fail: Arc<AtomicBool>,
}

struct OrderedReadListener {
    name: &'static str,
    events: Arc<Mutex<Vec<String>>>,
}

impl ReadListener<Value> for OrderedReadListener {
    fn invoke(&mut self, data: Value, _context: &AnalysisContext) -> Result<()> {
        self.events
            .lock()
            .expect("ordered listener lock")
            .push(format!("{}:{}", self.name, data.0));
        Ok(())
    }
}

impl Write for ToggleFacadeWrite {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        if self.fail.load(Ordering::SeqCst) {
            Err(io::Error::other("injected final encoding failure"))
        } else {
            Ok(buffer.len())
        }
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl ExcelRow for FallibleValue {
    fn schema() -> &'static [ExcelColumn] {
        Value::schema()
    }

    fn from_row(_row: &RowData) -> Result<Self> {
        Err(ExcelError::Unsupported("write-only test row".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        if self.fail {
            Err(ExcelError::Format("injected conversion failure".to_owned()))
        } else {
            Ok(vec![CellValue::String(self.value.to_owned())])
        }
    }
}

#[test]
fn writer_builder_excel_type_overrides_path_extension() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("values.data");

    EasyExcel::write::<Value>(&path)
        .excel_type(easyexcel_core::support::ExcelTypeEnum::Csv)
        .with_bom(false)
        .do_write(vec![Value("one".to_owned())])?;

    assert_eq!(fs::read_to_string(path)?, "Value\none\n");
    Ok(())
}

#[test]
fn reader_builder_register_read_listener_dispatches_in_registration_order() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("listeners.csv");
    fs::write(&path, "Value\none\ntwo\n")?;
    let events = Arc::new(Mutex::new(Vec::new()));

    EasyExcel::read::<Value, _>(
        &path,
        OrderedReadListener {
            name: "first",
            events: Arc::clone(&events),
        },
    )
    .register_read_listener(OrderedReadListener {
        name: "second",
        events: Arc::clone(&events),
    })
    .do_read()?;

    assert_eq!(
        *events.lock().expect("ordered listener lock"),
        vec![
            "first:one".to_owned(),
            "second:one".to_owned(),
            "first:two".to_owned(),
            "second:two".to_owned(),
        ]
    );
    Ok(())
}

fn write_minimal_template(path: &Path, shared_strings: &str, worksheet: &str) -> Result<()> {
    let file = fs::File::create(path)?;
    let mut archive = ZipWriter::new(file);
    archive
        .start_file("xl/sharedStrings.xml", SimpleFileOptions::default())
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    archive.write_all(shared_strings.as_bytes())?;
    archive
        .start_file("xl/worksheets/sheet1.xml", SimpleFileOptions::default())
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    archive.write_all(worksheet.as_bytes())?;
    archive
        .finish()
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    Ok(())
}

#[test]
fn facade_template_stream_factories_write_real_archives_and_close_owned_outputs() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("stream-template.xlsx");
    write_minimal_template(
        &template,
        "<sst><si><t>{name}</t></si></sst>",
        "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row></sheetData></worksheet>",
    )?;
    let bytes = fs::read(&template)?;

    let path_output = directory.path().join("reader-path.xlsx");
    EasyExcel::template_writer_from_reader(Cursor::new(bytes.clone()), &path_output)?.finish()?;
    assert!(fs::read(path_output)?.starts_with(b"PK"));

    let mut borrowed_path = Cursor::new(Vec::new());
    EasyExcel::template_writer_to_writer(&template, &mut borrowed_path)?.finish()?;
    assert!(borrowed_path.get_ref().starts_with(b"PK"));

    let mut borrowed_reader = Cursor::new(Vec::new());
    EasyExcel::template_writer_from_reader_to_writer(
        Cursor::new(bytes.clone()),
        &mut borrowed_reader,
    )?
    .finish()?;
    assert!(borrowed_reader.get_ref().starts_with(b"PK"));

    let path_stream = ExcelOutputStream::new(FacadeProbeWrite::default());
    let path_observer = path_stream.clone();
    EasyExcel::template_writer_to_output_stream(&template, path_stream)?.finish()?;
    assert!(path_observer.is_closed());

    let reader_stream = ExcelOutputStream::new(FacadeProbeWrite::default());
    let reader_observer = reader_stream.clone();
    EasyExcel::template_writer_from_reader_to_output_stream(Cursor::new(bytes), reader_stream)?
        .finish()?;
    assert!(reader_observer.is_closed());

    let missing = directory.path().join("missing-template.xlsx");
    assert!(
        EasyExcel::template_writer_from_reader(
            Cursor::new(b"invalid".to_vec()),
            directory.path().join("invalid-reader.xlsx")
        )
        .is_err()
    );
    let mut missing_borrowed = Cursor::new(Vec::new());
    assert!(EasyExcel::template_writer_to_writer(&missing, &mut missing_borrowed).is_err());
    assert!(
        EasyExcel::template_writer_from_reader_to_writer(
            Cursor::new(b"invalid".to_vec()),
            &mut missing_borrowed
        )
        .is_err()
    );
    assert!(
        EasyExcel::template_writer_to_output_stream(
            &missing,
            ExcelOutputStream::new(FacadeProbeWrite::default())
        )
        .is_err()
    );
    assert!(
        EasyExcel::template_writer_from_reader_to_output_stream(
            Cursor::new(b"invalid".to_vec()),
            ExcelOutputStream::new(FacadeProbeWrite::default())
        )
        .is_err()
    );
    Ok(())
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct ConverterRow {
    #[excel(name = "Value", index = 0)]
    value: String,
}

#[derive(Clone, Copy)]
struct PrefixConverter {
    prefix: &'static str,
    cell_type: CellDataType,
}

impl PrefixConverter {
    const fn string(prefix: &'static str) -> Self {
        Self {
            prefix,
            cell_type: CellDataType::String,
        }
    }
}

impl Converter<String> for PrefixConverter {
    fn support_excel_type(&self) -> CellDataType {
        self.cell_type
    }

    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<String> {
        Ok(format!(
            "{}:{}",
            self.prefix,
            context.cell().map_or_else(String::new, CellValue::as_text)
        ))
    }

    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<WriteCellData> {
        Ok(WriteCellData::from_string(format!(
            "{}:{}",
            self.prefix,
            context.value()
        )))
    }
}

#[derive(Default)]
struct FieldPrefixConverter;

impl Converter<String> for FieldPrefixConverter {
    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<String> {
        Ok(format!(
            "field:{}",
            context.cell().map_or_else(String::new, CellValue::as_text)
        ))
    }

    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<WriteCellData> {
        Ok(WriteCellData::from_string(format!(
            "field:{}",
            context.value()
        )))
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct FieldConverterRow {
    #[excel(name = "Value", index = 0, converter = FieldPrefixConverter)]
    value: String,
}

#[derive(Default)]
struct RejectingWriteConverter;

impl Converter<String> for RejectingWriteConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<WriteCellData> {
        if context.value() == "fail" {
            Err(ExcelError::Format("converter rejected value".to_owned()))
        } else {
            Ok(WriteCellData::from_string(context.value().clone()))
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct LocatedWriteFailureRow {
    #[excel(name = "Forced", index = 2)]
    forced: String,
    #[excel(name = "Late", order = 20)]
    late: String,
    #[excel(name = "Failing", order = 10, converter = RejectingWriteConverter)]
    failing: String,
}

struct WideCell(CellValue);

impl ExcelRow for WideCell {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] =
            &[ExcelColumn::new("value", "Value", Some(16_384), 0, None)];
        COLUMNS
    }

    fn from_row(_row: &RowData) -> Result<Self> {
        Err(ExcelError::Unsupported("write-only test row".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![self.0.clone()])
    }
}

struct SingleCell(CellValue);

impl ExcelRow for SingleCell {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Value", Some(0), 0, None)];
        COLUMNS
    }

    fn from_row(_row: &RowData) -> Result<Self> {
        Err(ExcelError::Unsupported("write-only test row".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![self.0.clone()])
    }
}

fn tiny_png() -> Vec<u8> {
    vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
        0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0xf0, 0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00,
        0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]
}

#[derive(Default)]
struct Listener(Vec<Value>);

struct FailingListener;

struct NoopWriteHandler;

#[derive(Clone, Default)]
struct DynamicListener(Arc<Mutex<Vec<DynamicRow>>>);

#[derive(Clone, Default)]
struct ConverterListener(Arc<Mutex<Vec<ConverterRow>>>);

impl WriteHandler for NoopWriteHandler {}

struct FailingFacadeWriteHandler {
    before_workbook: bool,
    before_cell: bool,
}

impl WriteHandler for FailingFacadeWriteHandler {
    fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        if self.before_workbook {
            Err(ExcelError::Format(
                "injected before-workbook failure".to_owned(),
            ))
        } else {
            Ok(())
        }
    }

    fn before_cell(&mut self, _context: &mut WriteCellContext) -> Result<()> {
        if self.before_cell {
            Err(ExcelError::Format(
                "injected before-cell failure".to_owned(),
            ))
        } else {
            Ok(())
        }
    }
}

impl ReadListener<Value> for Listener {
    fn invoke(&mut self, data: Value, _context: &AnalysisContext) -> Result<()> {
        self.0.push(data);
        Ok(())
    }
}

impl ReadListener<Value> for FailingListener {
    fn invoke_head(
        &mut self,
        _head: &std::collections::HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        Err(ExcelError::Format("injected listener failure".to_owned()))
    }

    fn invoke(&mut self, _data: Value, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }
}

impl ReadListener<DynamicRow> for DynamicListener {
    fn invoke(&mut self, data: DynamicRow, _context: &AnalysisContext) -> Result<()> {
        self.0.lock().expect("dynamic listener lock").push(data);
        Ok(())
    }
}

impl ReadListener<ConverterRow> for ConverterListener {
    fn invoke(&mut self, data: ConverterRow, _context: &AnalysisContext) -> Result<()> {
        self.0.lock().expect("converter listener lock").push(data);
        Ok(())
    }
}

#[test]
fn sheet_selector_inputs_map_indices_borrowed_and_owned_names() {
    assert_eq!(0_usize.into_sheet_selector(), SheetSelector::Index(0));
    assert_eq!(
        "Users".into_sheet_selector(),
        SheetSelector::Name("Users".to_owned())
    );
    assert_eq!(
        "Owned".to_owned().into_sheet_selector(),
        SheetSelector::Name("Owned".to_owned())
    );
    assert!(is_xls_path(Path::new("legacy.XLS")));
    assert!(!is_xls_path(Path::new("modern.xlsx")));
}

#[test]
#[allow(clippy::too_many_lines)]
fn factories_and_builder_options_match_java_style_chaining() {
    let read = EasyExcel::read::<Value, _>("input.xlsx", Listener::default())
        .sheet(2_usize)
        .all_sheets()
        .head_row_number(3)
        .ignore_empty_row(false)
        .auto_trim(false)
        .use_1904_windowing(true)
        .use_scientific_format(true)
        .use_scientific_format(false)
        .use_scientific_format(true)
        .locale(ExcelLocale::from_name("de-DE").expect("German locale"))
        .start_row(4)
        .end_row(8)
        .read_rows(5, 7)
        .header_alias("Source", "Value")
        .custom_object("event-context".to_owned())
        .read_cache(ReadCacheMode::Disk)
        .read_default_return(ReadDefaultReturn::ActualData)
        .extra_read(CellExtraType::Comment)
        .extra_read(CellExtraType::Merge)
        .password("read-secret")
        .charset("GBK");
    assert_eq!(read.path, PathBuf::from("input.xlsx"));
    assert_eq!(read.options.sheet, SheetSelector::All);
    assert_eq!(read.options.head_row_number, 3);
    assert!(!read.options.ignore_empty_row);
    assert!(!read.options.auto_trim);
    assert!(read.options.use_1904_windowing);
    assert_eq!(
        read.options.scientific_format,
        ScientificFormatMode::Scientific
    );
    assert_eq!(read.options.locale.language_tag(), "de_DE");
    assert_eq!(read.options.start_row, Some(5));
    assert_eq!(read.options.end_row, Some(7));
    assert_eq!(
        read.options
            .header_aliases
            .get("Source")
            .map(String::as_str),
        Some("Value")
    );
    assert_eq!(
        read.options
            .custom_object
            .as_ref()
            .and_then(|value| value.downcast_ref::<String>())
            .map(String::as_str),
        Some("event-context")
    );
    assert_eq!(
        read.options.read_default_return,
        ReadDefaultReturn::ActualData
    );
    assert_eq!(read.options.read_cache, ReadCacheMode::Disk);
    assert!(read.options.extra_read.contains(&CellExtraType::Comment));
    assert!(read.options.extra_read.contains(&CellExtraType::Merge));
    assert_eq!(read.options.password.as_deref(), Some("read-secret"));
    assert_eq!(read.options.charset.name(), "GBK");

    let sync = EasyExcel::read_sync::<Value>("sync.xlsx")
        .sheet("Values")
        .all_sheets()
        .head_row_number(2)
        .ignore_empty_row(false)
        .auto_trim(false)
        .use_1904_windowing(true)
        .use_scientific_format(true)
        .use_scientific_format(false)
        .use_scientific_format(true)
        .locale(ExcelLocale::from_name("zh-CN").expect("Chinese locale"))
        .start_row(3)
        .end_row(9)
        .read_rows(4, 6)
        .header_alias("Original", "Value")
        .custom_object(42_u32)
        .read_cache(ReadCacheMode::Memory)
        .read_default_return(ReadDefaultReturn::ReadCellData)
        .extra_read(CellExtraType::Hyperlink)
        .password("sync-secret")
        .charset(CsvCharset::new("UTF-16BE"));
    assert_eq!(sync.path, PathBuf::from("sync.xlsx"));
    assert_eq!(sync.options.sheet, SheetSelector::All);
    assert_eq!(sync.options.head_row_number, 2);
    assert!(!sync.options.ignore_empty_row);
    assert!(!sync.options.auto_trim);
    assert!(sync.options.use_1904_windowing);
    assert_eq!(
        sync.options.scientific_format,
        ScientificFormatMode::Scientific
    );
    assert_eq!(sync.options.locale.language_tag(), "zh_CN");
    assert_eq!(sync.options.start_row, Some(4));
    assert_eq!(sync.options.end_row, Some(6));
    assert_eq!(
        sync.options
            .header_aliases
            .get("Original")
            .map(String::as_str),
        Some("Value")
    );
    assert_eq!(
        sync.options
            .custom_object
            .as_ref()
            .and_then(|value| value.downcast_ref::<u32>()),
        Some(&42)
    );
    assert_eq!(
        sync.options.read_default_return,
        ReadDefaultReturn::ReadCellData
    );
    assert_eq!(sync.options.read_cache, ReadCacheMode::Memory);
    assert!(sync.options.extra_read.contains(&CellExtraType::Hyperlink));
    assert_eq!(sync.options.password.as_deref(), Some("sync-secret"));
    assert_eq!(sync.options.charset.name(), "UTF-16BE");

    let write = EasyExcel::write::<Value>("output.xlsx")
        .sheet("Values")
        .need_head(false)
        .freeze_head(true)
        .freeze_panes(2, 1)
        .include_column_indexes([2, 0])
        .include_column_field_names(["value"])
        .exclude_column_indexes([3])
        .exclude_column_field_names(["ignored".to_owned()])
        .order_by_include_column(true)
        .merge_cells(MergeRange::new(0, 0, 0, 1))
        .auto_width()
        .column_width(0, 24)
        .head_style(CellStyle::new().italic(true))
        .content_style(CellStyle::new().bold(true))
        .content_styles([CellStyle::new().wrap_text(true)])
        .loop_merge(LoopMergeStrategy::new(2, 1, 0).unwrap())
        .head([["Group", "Value"]])
        .password("write-secret")
        .charset("GBK")
        .with_bom(false)
        .register_write_handler(NoopWriteHandler)
        .constant_memory(true)
        .compress_temp_files(true);
    assert_eq!(write.path, PathBuf::from("output.xlsx"));
    assert_eq!(write.options.sheet_name, "Values");
    assert_eq!(write.options.sheet_index, None);
    assert!(!write.options.need_head);
    assert!(write.options.freeze_head);
    assert_eq!(write.options.freeze_panes, Some((2, 1)));
    assert_eq!(write.options.include_column_indexes, Some(vec![2, 0]));
    assert_eq!(
        write.options.include_column_field_names,
        Some(vec!["value".to_owned()])
    );
    assert_eq!(write.options.exclude_column_indexes, vec![3]);
    assert_eq!(
        write.options.exclude_column_field_names,
        vec!["ignored".to_owned()]
    );
    assert!(write.options.order_by_include_column);
    assert_eq!(
        write.options.merge_ranges,
        vec![MergeRange::new(0, 0, 0, 1)]
    );
    assert!(write.options.auto_width);
    assert_eq!(write.options.column_widths, vec![(0, 24)]);
    assert!(write.options.head_style.italic);
    assert_eq!(write.options.content_styles.len(), 1);
    assert!(write.options.content_styles[0].wrap_text);
    assert_eq!(write.options.loop_merges.len(), 1);
    assert_eq!(
        write.options.dynamic_head,
        Some(vec![vec!["Group".to_owned(), "Value".to_owned()]])
    );
    assert_eq!(write.handlers.len(), 1);
    assert!(write.options.constant_memory);
    assert!(write.options.compress_temp_files);
    assert_eq!(write.options.password.as_deref(), Some("write-secret"));
    assert_eq!(write.options.charset.name(), "GBK");
    assert!(!write.options.with_bom);

    let indexed_write = EasyExcel::write::<Value>("indexed.xlsx").sheet_index(4);
    assert_eq!(indexed_write.options.sheet_index, Some(4));
    assert_eq!(indexed_write.options.sheet_name, "4");
    let indexed_sheet = EasyExcel::writer_sheet_index::<Value>(5);
    assert_eq!(indexed_sheet.options().sheet_index, Some(5));
    assert_eq!(indexed_sheet.options().sheet_name, "5");

    let dynamic = EasyExcel::read_dynamic("dynamic.xlsx", DynamicListener::default());
    assert_eq!(dynamic.path, PathBuf::from("dynamic.xlsx"));
    assert_eq!(
        dynamic.options.read_default_return,
        ReadDefaultReturn::String
    );
    let dynamic_sync = EasyExcel::read_dynamic_sync("dynamic-sync.xlsx");
    assert_eq!(dynamic_sync.path, PathBuf::from("dynamic-sync.xlsx"));
}

#[test]
#[allow(clippy::too_many_lines)]
fn facade_reads_and_writes_java_style_dynamic_rows() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("dynamic.xlsx");
    let source = DynamicRow::new(BTreeMap::from([
        (0, DynamicValue::String("string19".to_owned())),
        (1, DynamicValue::ActualData(CellValue::Int(109))),
        (2, DynamicValue::Null),
        (3, DynamicValue::String("tail".to_owned())),
    ]));
    EasyExcel::write::<DynamicRow>(&path).do_write([source.clone()])?;

    let strings = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(
        strings[0].get(0),
        Some(&DynamicValue::String("string19".to_owned()))
    );
    assert_eq!(
        strings[0].get(1),
        Some(&DynamicValue::String("109".to_owned()))
    );
    assert_eq!(strings[0].get(2), Some(&DynamicValue::Null));

    let actual = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .read_default_return(ReadDefaultReturn::ActualData)
        .do_read_sync()?;
    let Some(DynamicValue::ActualData(number)) = actual[0].get(1) else {
        panic!("expected actual numeric cell");
    };
    assert_eq!(number.as_text(), "109");

    let listener = DynamicListener::default();
    let observed = Arc::clone(&listener.0);
    EasyExcel::read_dynamic(&path, listener)
        .head_row_number(0)
        .read_default_return(ReadDefaultReturn::ReadCellData)
        .do_read()?;
    let observed = observed.lock().expect("dynamic listener lock");
    let DynamicValue::ReadCellData(cell) = observed[0].get(3).expect("tail cell") else {
        panic!("expected read cell data");
    };
    assert_eq!(cell.data(), &CellValue::String("tail".to_owned()));

    let csv_without_head = directory.path().join("dynamic-no-head.csv");
    EasyExcel::write::<DynamicRow>(&csv_without_head)
        .with_bom(false)
        .do_write([source.clone()])?;
    let no_head_rows = EasyExcel::read_dynamic_sync(&csv_without_head)
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(
        no_head_rows[0].get(3),
        Some(&DynamicValue::String("tail".to_owned()))
    );
    assert!(matches!(
        EasyExcel::write::<DynamicRow>(directory.path().join("invalid-charset.csv"))
            .charset("not-a-charset")
            .do_write([source.clone()]),
        Err(ExcelError::Unsupported(_))
    ));
    assert!(
        EasyExcel::write::<DynamicRow>(directory.path().join("missing/dynamic.csv"))
            .do_write([source.clone()])
            .is_err()
    );

    let csv = directory.path().join("dynamic.csv");
    EasyExcel::write::<DynamicRow>(&csv)
        .head([["Text"], ["Number"], ["Empty"], ["Tail"]])
        .with_bom(false)
        .do_write([source])?;
    let csv_rows = EasyExcel::read_dynamic_sync(&csv).do_read_sync()?;
    assert_eq!(
        csv_rows[0].get(0),
        Some(&DynamicValue::String("string19".to_owned()))
    );
    assert_eq!(
        csv_rows[0].get(1),
        Some(&DynamicValue::String("109".to_owned()))
    );

    let filter_source = DynamicRow::new(BTreeMap::from([
        (0, DynamicValue::String("A".to_owned())),
        (1, DynamicValue::String("B".to_owned())),
        (2, DynamicValue::String("C".to_owned())),
    ]));
    let filtered = directory.path().join("dynamic-filtered.xlsx");
    EasyExcel::write::<DynamicRow>(&filtered)
        .include_column_indexes([2, 0])
        .exclude_column_indexes([2])
        .order_by_include_column(true)
        .do_write([filter_source.clone()])?;
    assert_eq!(
        EasyExcel::read_dynamic_sync(&filtered)
            .head_row_number(0)
            .do_read_sync()?[0]
            .get(0),
        Some(&DynamicValue::String("A".to_owned()))
    );

    EasyExcel::write::<DynamicRow>(directory.path().join("dynamic-ordered.xlsx"))
        .order_by_include_column(true)
        .do_write([filter_source.clone()])?;
    EasyExcel::write::<DynamicRow>(directory.path().join("dynamic-field-filter.xlsx"))
        .include_column_field_names(["unknown"])
        .do_write([filter_source.clone()])?;
    EasyExcel::write::<DynamicRow>(directory.path().join("dynamic-index-include.xlsx"))
        .include_column_indexes([1])
        .do_write([filter_source])?;
    EasyExcel::write::<Value>(directory.path().join("typed-field-include.xlsx"))
        .include_column_field_names(["value"])
        .do_write([Value("included".to_owned())])?;
    Ok(())
}

#[test]
fn facade_applies_1904_windowing_to_numeric_date_converters() -> Result<()> {
    #[derive(Debug, PartialEq, ExcelRow)]
    struct NumericDates {
        #[excel(index = 0, use_1904_windowing = true)]
        date: NaiveDate,
        #[excel(index = 1, use_1904_windowing = false)]
        datetime: chrono::NaiveDateTime,
    }

    let directory = tempdir()?;
    let path = directory.path().join("numeric-date-1904.xlsx");
    let source = DynamicRow::new(BTreeMap::from([
        (0, DynamicValue::ActualData(CellValue::Int(0))),
        (1, DynamicValue::ActualData(CellValue::Float(1.5))),
    ]));
    EasyExcel::write::<DynamicRow>(&path).do_write([source])?;

    let rows = EasyExcel::read_sync::<NumericDates>(&path)
        .head_row_number(0)
        .use_1904_windowing(true)
        .do_read_sync()?;
    assert_eq!(
        rows,
        vec![NumericDates {
            date: NaiveDate::from_ymd_opt(1904, 1, 1).expect("date"),
            datetime: NaiveDate::from_ymd_opt(1900, 1, 1)
                .expect("date")
                .and_hms_opt(12, 0, 0)
                .expect("time"),
        }]
    );
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn facade_executes_event_sync_and_iterator_workflows() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("values.xlsx");
    let rows = vec![Value("one".to_owned()), Value("two".to_owned())];
    EasyExcel::write::<Value>(&path)
        .sheet("Values")
        .freeze_head(true)
        .do_write_iter(rows.clone())?;

    let actual = EasyExcel::read_sync::<Value>(&path)
        .sheet("Values".to_owned())
        .do_read_sync()?;
    assert_eq!(actual, rows);

    let csv = directory.path().join("values.CSV");
    EasyExcel::write::<Value>(&csv).do_write(rows.clone())?;
    assert_eq!(EasyExcel::read_sync::<Value>(&csv).do_read_sync()?, rows);
    EasyExcel::read::<Value, _>(&csv, Listener::default())
        .sheet("CsvSheet")
        .do_read()?;

    let gbk_csv = directory.path().join("values-gbk.csv");
    let chinese = vec![Value("姓名".repeat(5_000))];
    EasyExcel::write::<Value>(&gbk_csv)
        .charset("GBK")
        .with_bom(false)
        .do_write(chinese.clone())?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&gbk_csv)
            .charset("gbk")
            .do_read_sync()?,
        chinese
    );
    EasyExcel::read::<Value, _>(&gbk_csv, Listener::default())
        .charset("GBK")
        .do_read()?;
    assert!(matches!(
        EasyExcel::write::<Value>(directory.path().join("protected.csv"))
            .password("secret")
            .do_write(rows.clone()),
        Err(ExcelError::Unsupported(_))
    ));

    let encrypted = directory.path().join("protected.xlsx");
    EasyExcel::write::<Value>(&encrypted)
        .password("123456")
        .do_write(rows.clone())?;
    assert_eq!(
        &fs::read(&encrypted)?[..8],
        &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
    );
    assert_eq!(
        EasyExcel::read_sync::<Value>(&encrypted)
            .password("123456")
            .do_read_sync()?,
        rows
    );
    assert!(
        EasyExcel::read_sync::<Value>(&encrypted)
            .password("wrong")
            .do_read_sync()
            .is_err()
    );
    assert!(
        EasyExcel::read_sync::<Value>(&encrypted)
            .do_read_sync()
            .is_err()
    );
    let invalid_encrypted = directory.path().join("invalid-encrypted.xlsx");
    fs::write(
        &invalid_encrypted,
        [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
    )?;
    assert!(
        EasyExcel::read_sync::<Value>(&invalid_encrypted)
            .password("123456")
            .do_read_sync()
            .is_err()
    );
    assert_eq!(
        EasyExcel::read_sync::<Value>(&path)
            .password("ignored-for-plain-xlsx")
            .sheet("Values")
            .do_read_sync()?,
        rows
    );
    assert!(
        EasyExcel::read_sync::<Value>(&path)
            .sheet(99_usize)
            .do_read_sync()
            .is_err()
    );
    assert!(
        EasyExcel::read::<Value, _>(&path, FailingListener)
            .do_read()
            .is_err()
    );

    EasyExcel::read::<Value, _>(&path, Listener::default())
        .all_sheets()
        .do_read()?;

    let no_head = directory.path().join("no-head.xlsx");
    EasyExcel::write::<Value>(&no_head)
        .need_head(false)
        .constant_memory(true)
        .do_write(rows.clone())?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&no_head)
            .head_row_number(0)
            .do_read_sync()?
            .len(),
        2
    );

    let multi = directory.path().join("multi.xlsx");
    let first = EasyExcel::writer_sheet::<Value>("First").freeze_head(true);
    let second = EasyExcel::writer_sheet::<Value>("Second")
        .need_head(false)
        .constant_memory(true);
    let mut writer = EasyExcel::write::<Value>(&multi)
        .register_write_handler(NoopWriteHandler)
        .build();
    writer
        .write(vec![Value("first".to_owned())], &first)?
        .write(vec![Value("second".to_owned())], &second)?;
    writer.finish()?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&multi)
            .sheet("First")
            .do_read_sync()?,
        vec![Value("first".to_owned())]
    );
    assert_eq!(
        EasyExcel::read_sync::<Value>(&multi)
            .sheet("Second")
            .head_row_number(0)
            .do_read_sync()?,
        vec![Value("second".to_owned())]
    );

    let encrypted_multi = directory.path().join("encrypted-multi.xlsx");
    let mut encrypted_writer = EasyExcel::write::<Value>(&encrypted_multi)
        .password("stateful")
        .build();
    encrypted_writer.write(rows.clone(), &first)?.finish()?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&encrypted_multi)
            .password("stateful")
            .sheet("First")
            .do_read_sync()?,
        rows
    );

    let template = directory.path().join("template.xlsx");
    let filled = directory.path().join("filled.xlsx");
    EasyExcel::write::<Value>(&template)
        .need_head(false)
        .do_write(vec![Value("Hello {name}".to_owned())])?;
    EasyExcel::fill_template(
        &template,
        &filled,
        &TemplateData::new().with("name", "Rust"),
    )?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&filled)
            .head_row_number(0)
            .do_read_sync()?,
        vec![Value("Hello Rust".to_owned())]
    );

    let typed_template = directory.path().join("typed-template.xlsx");
    let typed_filled = directory.path().join("typed-filled.xlsx");
    EasyExcel::write::<Value>(&typed_template)
        .need_head(false)
        .do_write(vec![Value("{number}".to_owned())])?;
    EasyExcel::fill_template(
        &typed_template,
        &typed_filled,
        &TemplateData::new().with("number", BigDecimal::from(42)),
    )?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&typed_filled)
            .head_row_number(0)
            .do_read_sync()?,
        vec![Value("42".to_owned())]
    );

    let list_template = directory.path().join("list-template.xlsx");
    let list_filled = directory.path().join("list-filled.xlsx");
    EasyExcel::write::<Value>(&list_template)
        .need_head(false)
        .do_write(vec![Value("{.name}".to_owned())])?;
    EasyExcel::fill_template_list(
        &list_template,
        &list_filled,
        &FillWrapper::new([
            TemplateData::new().with("name", "one"),
            TemplateData::new().with("name", "two"),
        ]),
        FillConfig::new(),
    )?;
    assert_eq!(
        EasyExcel::read_sync::<Value>(&list_filled)
            .head_row_number(0)
            .do_read_sync()?,
        vec![Value("one".to_owned()), Value("two".to_owned())]
    );

    let repeated_filled = directory.path().join("list-repeated-filled.xlsx");
    let mut template_writer =
        EasyExcel::template_writer(list_template.clone(), repeated_filled.clone())?;
    template_writer
        .fill(&TemplateData::new())?
        .fill_list(
            &FillWrapper::new([TemplateData::new().with("name", "first")]),
            FillConfig::new(),
        )?
        .fill_list(
            &FillWrapper::new([TemplateData::new().with("name", "second")]),
            FillConfig::new(),
        )?
        .fill_list(&FillWrapper::default(), FillConfig::new())?
        .write_rows([vec![CellValue::String("summary".to_owned())]])?;
    template_writer.fill_list(
        &FillWrapper::new([TemplateData::new().with("name", "horizontal")]),
        FillConfig::new().direction(FillDirection::Horizontal),
    )?;
    template_writer.finish()?;
    template_writer.finish()?;
    assert!(template_writer.fill(&TemplateData::new()).is_err());
    assert!(
        template_writer
            .write_rows([Vec::<CellValue>::new()])
            .is_err()
    );
    assert!(
        template_writer
            .fill_list(&FillWrapper::default(), FillConfig::new())
            .is_err()
    );
    assert_eq!(
        EasyExcel::read_sync::<Value>(&repeated_filled)
            .head_row_number(0)
            .do_read_sync()?,
        vec![
            Value("first".to_owned()),
            Value("second".to_owned()),
            Value("summary".to_owned())
        ]
    );
    assert!(
        EasyExcel::template_writer(
            directory.path().join("missing-template.xlsx"),
            directory.path().join("missing-output.xlsx"),
        )
        .is_err()
    );
    assert!(
        EasyExcel::fill_template(
            directory.path().join("missing-template.xlsx"),
            directory.path().join("missing-output.xlsx"),
            &TemplateData::new(),
        )
        .is_err()
    );
    assert!(
        EasyExcel::fill_template_list(
            directory.path().join("missing-template.xlsx"),
            directory.path().join("missing-output.xlsx"),
            &FillWrapper::default(),
            FillConfig::new(),
        )
        .is_err()
    );
    assert!(
        EasyExcel::fill_template_list(
            &list_template,
            directory.path().join("missing/template-output.xlsx"),
            &FillWrapper::default(),
            FillConfig::new(),
        )
        .is_err()
    );

    let malformed_template = directory.path().join("malformed-template.xlsx");
    let malformed_output = directory.path().join("malformed-output.xlsx");
    write_minimal_template(
        &malformed_template,
        "<sst><si><t>{.name}</t></si><si><t</si><si><t>missing</si></sst>",
        concat!(
            "<worksheet><sheetData><row r=\"1\">",
            "<c t=\"s\"></c><c t=\"s\"><v>broken</c><c t=\"s\"><v>9</v></c>",
            "<c t=\"inlineStr\"><is><t</is></c>",
            "<c t=\"inlineStr\"><is><t>missing</is></c>",
            "<c r=\"A1\" t=\"s\"><v>0</v></c>",
            "<c r=\"B1\"><v>{.name}</v></c>",
            "</row></sheetData></worksheet>"
        ),
    )?;
    EasyExcel::fill_template_list(
        &malformed_template,
        &malformed_output,
        &FillWrapper::new([TemplateData::new().with("name", "covered")]),
        FillConfig::new(),
    )?;

    let untyped_template = directory.path().join("untyped-template.xlsx");
    write_minimal_template(
        &untyped_template,
        "<sst></sst>",
        concat!(
            "<worksheet><sheetData><row r=\"1\">",
            "<c r=\"A1\"><v>{.name}</v></c>",
            "</row></sheetData></worksheet>"
        ),
    )?;
    EasyExcel::fill_template_list(
        &untyped_template,
        directory.path().join("untyped-output.xlsx"),
        &FillWrapper::new([TemplateData::new().with("name", "covered")]),
        FillConfig::new(),
    )?;
    Ok(())
}

#[test]
fn facade_builds_stateful_gbk_csv_and_appends_without_repeating_head() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("stateful.csv");
    let sheet = EasyExcel::writer_sheet::<Value>("Values");
    let mut writer = EasyExcel::write::<Value>(&path)
        .charset("GBK")
        .with_bom(false)
        .build();
    writer
        .write(vec![Value("第一批".to_owned())], &sheet)?
        .write(vec![Value("第二批".to_owned())], &sheet)?;
    writer.finish()?;
    writer.finish()?;
    let mut empty_writer = EasyExcel::write::<Value>(directory.path().join("empty.csv")).build();
    empty_writer.finish()?;

    assert_eq!(
        EasyExcel::read_sync::<Value>(&path)
            .charset("gbk")
            .do_read_sync()?,
        vec![Value("第一批".to_owned()), Value("第二批".to_owned())]
    );
    Ok(())
}

#[test]
fn facade_csv_stream_writer_propagates_validation_and_io_failures() {
    let mut stream_options = WriteOptions {
        with_bom: false,
        ..WriteOptions::default()
    };
    assert!(
        write_csv_to_writer::<Value, _, _>(
            Path::new("stream.csv"),
            Cursor::new(Vec::new()),
            &stream_options,
            [Value("streamed".to_owned())],
            &mut [],
        )
        .is_ok()
    );
    assert!(matches!(
        write_csv_to_writer::<Value, _, _>(
            Path::new("stream.csv"),
            Cursor::new(Vec::new().into_boxed_slice()),
            &stream_options,
            [Value("output failure".to_owned())],
            &mut [],
        ),
        Err(ExcelError::Io(_) | ExcelError::Format(_))
    ));
    stream_options.charset = CsvCharset::new("not-a-real-charset");
    assert!(matches!(
        write_csv_to_writer::<Value, _, _>(
            Path::new("stream.csv"),
            Cursor::new(Vec::new()),
            &stream_options,
            [Value("ignored".to_owned())],
            &mut [],
        ),
        Err(ExcelError::Unsupported(_))
    ));
    assert!(
        write_csv_to_writer::<Value, _, _>(
            Path::new("stream.csv"),
            FacadeProbeWrite {
                fail_write: true,
                ..FacadeProbeWrite::default()
            },
            &WriteOptions::default(),
            [Value("bom failure".to_owned())],
            &mut [],
        )
        .is_err()
    );
    for fail_flush_at in [0, 1] {
        assert!(matches!(
            write_csv_to_writer::<Value, _, _>(
                Path::new("stream.csv"),
                FacadeProbeWrite {
                    fail_flush_at: Some(fail_flush_at),
                    ..FacadeProbeWrite::default()
                },
                &WriteOptions {
                    with_bom: false,
                    ..WriteOptions::default()
                },
                [Value("flush failure".to_owned())],
                &mut [],
            ),
            Err(ExcelError::Io(_) | ExcelError::Format(_))
        ));
    }

    let mut incomplete =
        CsvEncodingWriter::with_charset(FacadeProbeWrite::default(), &CsvCharset::new("UTF-8"))
            .expect("UTF-8 transcoder");
    assert!(matches!(
        CsvEncodingWriter::with_charset(
            FacadeProbeWrite::default(),
            &CsvCharset::new("not-a-real-charset"),
        ),
        Err(ExcelError::Unsupported(_))
    ));
    incomplete.write_all(&[0xE2]).expect("partial UTF-8 chunk");
    assert_eq!(
        incomplete
            .finish()
            .expect_err("incomplete UTF-8 fails")
            .kind(),
        io::ErrorKind::InvalidData
    );

    let fail = Arc::new(AtomicBool::new(false));
    let mut finalizing = CsvEncodingWriter::with_charset(
        ToggleFacadeWrite {
            fail: Arc::clone(&fail),
        },
        &CsvCharset::new("ISO-2022-JP"),
    )
    .expect("ISO-2022-JP transcoder");
    finalizing
        .write_all("日本".as_bytes())
        .expect("initial encoded bytes");
    fail.store(true, Ordering::SeqCst);
    assert!(finalizing.finish().is_err());
}

#[test]
#[allow(clippy::too_many_lines)]
fn facade_borrowed_xlsx_stream_is_real_and_remains_caller_owned() -> Result<()> {
    let mut output = FacadeProbeWrite::default();
    EasyExcel::write::<Value>("response.xlsx")
        .sheet("Values")
        .to_writer(&mut output)
        .do_write([Value("streamed".to_owned())])?;
    assert!(output.bytes.starts_with(b"PK"));
    output.write_all(b"caller-still-owns-stream")?;
    assert!(output.bytes.ends_with(b"caller-still-owns-stream"));

    let mut encrypted = FacadeProbeWrite::default();
    EasyExcel::write::<Value>("response.xlsx")
        .password("123456")
        .to_writer(&mut encrypted)
        .do_write([Value("secret".to_owned())])?;
    assert!(encrypted.bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]));

    let mut csv = FacadeProbeWrite::default();
    EasyExcel::write::<Value>("response.csv")
        .with_bom(false)
        .to_writer(&mut csv)
        .do_write([Value("csv-stream".to_owned())])?;
    assert_eq!(csv.bytes, b"Value\ncsv-stream\n");

    for charset in ["UTF-16LE", "UTF-16BE"] {
        let mut encoded = FacadeProbeWrite::default();
        EasyExcel::write::<Value>("response.csv")
            .charset(charset)
            .to_writer(&mut encoded)
            .do_write([Value("encoded".to_owned())])?;
        assert!(!encoded.bytes.is_empty());
    }

    let mut invalid_csv = FacadeProbeWrite::default();
    assert!(matches!(
        EasyExcel::write::<Value>("response.csv")
            .charset("not-a-charset")
            .to_writer(&mut invalid_csv)
            .do_write([Value("invalid".to_owned())]),
        Err(ExcelError::Unsupported(_))
    ));

    for mut output in [
        FacadeProbeWrite {
            fail_write: true,
            ..FacadeProbeWrite::default()
        },
        FacadeProbeWrite {
            fail_flush: true,
            ..FacadeProbeWrite::default()
        },
    ] {
        assert!(
            EasyExcel::write::<Value>("response.csv")
                .with_bom(false)
                .to_writer(&mut output)
                .do_write([Value("failure".to_owned())])
                .is_err()
        );
    }
    for mut output in [
        FacadeProbeWrite {
            fail_write: true,
            ..FacadeProbeWrite::default()
        },
        FacadeProbeWrite {
            fail_flush: true,
            ..FacadeProbeWrite::default()
        },
    ] {
        assert!(
            EasyExcel::write::<Value>("response.xlsx")
                .to_writer(&mut output)
                .do_write([Value("failure".to_owned())])
                .is_err()
        );
    }
    let mut encrypted_failure = FacadeProbeWrite {
        fail_write: true,
        ..FacadeProbeWrite::default()
    };
    assert!(
        EasyExcel::write::<Value>("response.xlsx")
            .password("123456")
            .to_writer(&mut encrypted_failure)
            .do_write([Value("failure".to_owned())])
            .is_err()
    );
    for (before_workbook, before_cell) in [(true, false), (false, true)] {
        let mut output = FacadeProbeWrite::default();
        assert!(
            EasyExcel::write::<Value>("response.xlsx")
                .register_write_handler(FailingFacadeWriteHandler {
                    before_workbook,
                    before_cell,
                })
                .to_writer(&mut output)
                .do_write([Value("failure".to_owned())])
                .is_err()
        );
    }

    let mut xls_stream = FacadeProbeWrite::default();
    EasyExcel::write::<Value>("response.xls")
        .to_writer(&mut xls_stream)
        .do_write([Value("biff8".to_owned())])?;
    // OLE/CFB compound-document signature (D0 CF 11 E0).
    assert!(xls_stream.bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]));
    // Phase 5.3: XLS password is now supported via BIFF8 RC4
    assert!(
        EasyExcel::write::<Value>("response.xls")
            .password("123456")
            .to_writer(&mut FacadeProbeWrite::default())
            .do_write([Value("encrypted".to_owned())])
            .is_ok()
    );
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn facade_owned_stream_matches_close_and_exception_finish_semantics() -> Result<()> {
    let xlsx = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_xlsx = xlsx.clone();
    let sheet = EasyExcel::writer_sheet::<Value>("Values");
    let mut writer = EasyExcel::write::<Value>("response.xlsx")
        .auto_close_stream(false)
        .to_output_stream(xlsx)
        .build();
    writer.write([Value("one".to_owned())], &sheet)?;
    writer.write([Value("two".to_owned())], &sheet)?;
    writer.finish()?;
    writer.finish()?;
    assert!(writer.is_finished());
    assert!(!inspect_xlsx.is_closed());
    assert!(
        inspect_xlsx
            .with_inner(|output| output.bytes.starts_with(b"PK"))
            .unwrap_or(false)
    );

    let csv = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_csv = csv.clone();
    let csv_sheet = EasyExcel::writer_sheet::<Value>("Values");
    let mut csv_writer = EasyExcel::write::<Value>("response.csv")
        .with_bom(false)
        .auto_close_stream(false)
        .to_output_stream(csv)
        .build();
    csv_writer.write([Value("one".to_owned())], &csv_sheet)?;
    csv_writer.write([Value("two".to_owned())], &csv_sheet)?;
    csv_writer.finish()?;
    assert_eq!(
        inspect_csv.with_inner(|output| output.bytes.clone()),
        Some(b"Value\none\ntwo\n".to_vec())
    );
    let mut invalid_csv_writer = EasyExcel::write::<Value>("response.csv")
        .charset("not-a-charset")
        .to_output_stream(ExcelOutputStream::new(FacadeProbeWrite::default()))
        .build();
    assert!(matches!(
        invalid_csv_writer.finish(),
        Err(ExcelError::Unsupported(_))
    ));

    let discarded = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_discarded = discarded.clone();
    let mut discarded_writer = EasyExcel::write::<FallibleValue>("response.xlsx")
        .auto_close_stream(false)
        .to_output_stream(discarded)
        .build();
    let fallible_sheet = EasyExcel::writer_sheet::<FallibleValue>("Values");
    discarded_writer.write(
        [FallibleValue {
            value: "kept-in-workbook",
            fail: false,
        }],
        &fallible_sheet,
    )?;
    assert!(
        discarded_writer
            .write(
                [FallibleValue {
                    value: "fail",
                    fail: true,
                }],
                &fallible_sheet,
            )
            .is_err()
    );
    discarded_writer.finish_on_exception()?;
    assert_eq!(
        inspect_discarded.with_inner(|output| output.bytes.len()),
        Some(0)
    );

    let emitted = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_emitted = emitted.clone();
    let mut emitted_writer = EasyExcel::write::<FallibleValue>("response.xlsx")
        .auto_close_stream(false)
        .write_excel_on_exception(true)
        .to_output_stream(emitted)
        .build();
    emitted_writer.write(
        [FallibleValue {
            value: "emitted",
            fail: false,
        }],
        &fallible_sheet,
    )?;
    emitted_writer.finish_on_exception()?;
    assert!(
        inspect_emitted
            .with_inner(|output| output.bytes.starts_with(b"PK"))
            .unwrap_or(false)
    );

    let discarded_csv = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_discarded_csv = discarded_csv.clone();
    let mut discarded_csv_writer = EasyExcel::write::<FallibleValue>("response.csv")
        .with_bom(false)
        .auto_close_stream(false)
        .to_output_stream(discarded_csv)
        .build();
    discarded_csv_writer.write(
        [FallibleValue {
            value: "discarded",
            fail: false,
        }],
        &fallible_sheet,
    )?;
    discarded_csv_writer.finish_on_exception()?;
    assert_eq!(
        inspect_discarded_csv.with_inner(|output| output.bytes.len()),
        Some(0)
    );

    let emitted_csv = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_emitted_csv = emitted_csv.clone();
    let mut emitted_csv_writer = EasyExcel::write::<FallibleValue>("response.csv")
        .with_bom(false)
        .auto_close_stream(false)
        .write_excel_on_exception(true)
        .to_output_stream(emitted_csv)
        .build();
    emitted_csv_writer.write(
        [FallibleValue {
            value: "emitted",
            fail: false,
        }],
        &fallible_sheet,
    )?;
    emitted_csv_writer.finish_on_exception()?;
    assert_eq!(
        inspect_emitted_csv.with_inner(|output| output.bytes.clone()),
        Some(b"Value\nemitted\n".to_vec())
    );

    for output in [
        FacadeProbeWrite {
            fail_write: true,
            ..FacadeProbeWrite::default()
        },
        FacadeProbeWrite {
            fail_flush: true,
            ..FacadeProbeWrite::default()
        },
    ] {
        let mut failed_commit = EasyExcel::write::<Value>("response.csv")
            .with_bom(false)
            .auto_close_stream(false)
            .to_output_stream(ExcelOutputStream::new(output))
            .build();
        failed_commit.write([Value("failure".to_owned())], &csv_sheet)?;
        assert!(failed_commit.finish().is_err());
    }

    let one_shot = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_one_shot = one_shot.clone();
    assert!(
        EasyExcel::write::<FallibleValue>("response.xlsx")
            .auto_close_stream(false)
            .write_excel_on_exception(true)
            .to_output_stream(one_shot)
            .do_write([
                FallibleValue {
                    value: "emitted-before-error",
                    fail: false,
                },
                FallibleValue {
                    value: "error",
                    fail: true,
                },
            ])
            .is_err()
    );
    assert!(
        inspect_one_shot
            .with_inner(|output| output.bytes.starts_with(b"PK"))
            .unwrap_or(false)
    );

    let cleanup_failure = ExcelOutputStream::new(FacadeProbeWrite {
        fail_write: true,
        ..FacadeProbeWrite::default()
    });
    assert!(
        EasyExcel::write::<FallibleValue>("response.xlsx")
            .auto_close_stream(false)
            .write_excel_on_exception(true)
            .to_output_stream(cleanup_failure)
            .do_write([
                FallibleValue {
                    value: "emitted-before-error",
                    fail: false,
                },
                FallibleValue {
                    value: "error",
                    fail: true,
                },
            ])
            .is_err()
    );

    let mut invalid_encrypted = EasyExcel::write::<Value>("response.xlsx")
        .password("123456")
        .auto_close_stream(false)
        .to_output_stream(ExcelOutputStream::new(FacadeProbeWrite::default()))
        .build();
    invalid_encrypted
        .workbook_mut()
        .add_worksheet()
        .set_name("Duplicate")
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    invalid_encrypted
        .workbook_mut()
        .add_worksheet()
        .set_name("Duplicate")
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    assert!(invalid_encrypted.finish().is_err());

    let closed = ExcelOutputStream::new(FacadeProbeWrite::default());
    let inspect_closed = closed.clone();
    EasyExcel::write::<FallibleValue>("response.xlsx")
        .to_output_stream(closed)
        .do_write([
            FallibleValue {
                value: "closed-one",
                fail: false,
            },
            FallibleValue {
                value: "closed-two",
                fail: false,
            },
        ])?;
    assert!(inspect_closed.is_closed());
    inspect_closed.close()?;
    let mut closed_writer = inspect_closed.clone();
    assert!(closed_writer.write_all(b"rejected").is_err());
    assert!(closed_writer.flush().is_err());

    let failed_close = ExcelOutputStream::new(FacadeProbeWrite {
        fail_flush: true,
        ..FacadeProbeWrite::default()
    });
    assert!(failed_close.close().is_err());

    let poisoned_close = ExcelOutputStream::new(FacadeProbeWrite::default());
    let panic_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        poisoned_close.with_inner(|_| panic!("poison facade output lock"));
    }));
    assert!(panic_result.is_err());
    assert!(poisoned_close.close().is_err());
    let mut poisoned_writer = poisoned_close.clone();
    assert!(poisoned_writer.write_all(b"rejected").is_err());
    assert!(poisoned_writer.flush().is_err());
    Ok(())
}

#[test]
fn facade_propagates_read_sync_and_write_failures() {
    let missing = PathBuf::from("target/does-not-exist/easyexcel.xlsx");
    assert!(
        EasyExcel::read::<Value, _>(&missing, Listener::default())
            .do_read()
            .is_err()
    );
    assert!(
        EasyExcel::read_sync::<Value>(&missing)
            .do_read_sync()
            .is_err()
    );
    assert!(
        EasyExcel::read_sync::<Value>("target/does-not-exist/easyexcel.csv")
            .do_read_sync()
            .is_err()
    );
    assert!(
        EasyExcel::read::<Value, _>("target/does-not-exist/easyexcel.xls", Listener::default())
            .do_read()
            .is_err()
    );
    assert!(
        EasyExcel::read_sync::<Value>("target/does-not-exist/easyexcel.xls")
            .do_read_sync()
            .is_err()
    );
    assert!(
        EasyExcel::write::<Value>("target/does-not-exist/output.xlsx")
            .do_write(Vec::new())
            .is_err()
    );
    assert!(
        EasyExcel::write::<Value>("target/does-not-exist/output.csv")
            .do_write(Vec::new())
            .is_err()
    );
    assert!(
        EasyExcel::write::<Value>("target/does-not-exist/encrypted.xlsx")
            .password("123456")
            .do_write(Vec::new())
            .is_err()
    );

    let directory = tempdir().expect("temporary directory");
    // Minimal BIFF8: empty .xls write succeeds (password / template still Unsupported).
    let xls_empty = directory.path().join("empty.xls");
    EasyExcel::write::<Value>(&xls_empty)
        .do_write(Vec::<Value>::new())
        .expect("empty BIFF8 write");
    assert!(xls_empty.exists());
    // Phase 5.3: XLS password is now supported via BIFF8 RC4
    assert!(
        EasyExcel::write::<Value>(directory.path().join("encrypted.xls"))
            .password("123456")
            .do_write(Vec::new())
            .is_ok()
    );

    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    for (index, value) in [
        CellValue::Empty,
        CellValue::String("text".to_owned()),
        CellValue::Error("#DIV/0!".to_owned()),
        CellValue::Bool(true),
        CellValue::Int(1),
        CellValue::Int(i64::MAX),
        CellValue::Float(1.25),
        CellValue::Date(date),
        CellValue::DateTime(date.and_hms_opt(12, 34, 56).expect("valid time")),
        CellValue::Formula("1+1".to_owned()),
        CellValue::Hyperlink {
            url: "https://www.rust-lang.org".to_owned(),
            text: "Rust".to_owned(),
        },
        CellValue::Comment {
            value: Box::new(CellValue::String("annotated".to_owned())),
            text: "cell note".to_owned(),
        },
        CellValue::Image(vec![1, 2, 3]),
        CellValue::Image(tiny_png()),
    ]
    .into_iter()
    .enumerate()
    {
        assert!(
            EasyExcel::write::<WideCell>(directory.path().join(format!("wide-cell-{index}.xlsx")))
                .need_head(false)
                .do_write([WideCell(value)])
                .is_err()
        );
    }
    assert!(
        EasyExcel::write::<SingleCell>(directory.path().join("oversized-comment.xlsx"))
            .need_head(false)
            .do_write([SingleCell(CellValue::Comment {
                value: Box::new(CellValue::String("annotated".to_owned())),
                text: "x".repeat(32_768),
            })])
            .is_err()
    );
}

#[test]
fn collecting_listener_appends_rows() -> Result<()> {
    let mut listener = CollectListener(Vec::new());
    listener.invoke(
        Value("value".to_owned()),
        &AnalysisContext::new("Sheet1", 0, 1),
    )?;
    assert_eq!(listener.0, vec![Value("value".to_owned())]);
    Ok(())
}

#[test]
fn registered_converter_runs_in_sync_and_event_read_paths() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("registered-read.xlsx");
    EasyExcel::write::<ConverterRow>(&path).do_write([ConverterRow {
        value: "source".to_owned(),
    }])?;

    let rows = EasyExcel::read_sync::<ConverterRow>(&path)
        .register_converter::<String, _>(PrefixConverter::string("sync"))
        .do_read_sync()?;
    assert_eq!(rows[0].value, "sync:source");

    let probe = ConverterListener::default();
    let observed = Arc::clone(&probe.0);
    EasyExcel::read::<ConverterRow, _>(&path, probe)
        .register_converter::<String, _>(PrefixConverter::string("event"))
        .do_read()?;
    assert_eq!(
        observed.lock().expect("converter listener lock")[0].value,
        "event:source"
    );

    let fallback = EasyExcel::read_sync::<ConverterRow>(&path)
        .register_converter::<String, _>(PrefixConverter {
            prefix: "wrong-cell-type",
            cell_type: CellDataType::Boolean,
        })
        .do_read_sync()?;
    assert_eq!(fallback[0].value, "source");
    Ok(())
}

#[test]
fn registered_write_converter_uses_latest_registration_and_field_precedence() -> Result<()> {
    struct OriginalValueProbe(
        Arc<Mutex<Vec<(Option<CellValue>, Option<&'static str>, CellValue)>>>,
    );

    impl WriteHandler for OriginalValueProbe {
        fn after_cell_data_converted(&mut self, context: &WriteCellContext) -> Result<()> {
            if !context.is_head {
                self.0
                    .lock()
                    .map_err(|_| ExcelError::Format("converter probe poisoned".to_owned()))?
                    .push((
                        context.original_value.clone(),
                        context.original_field_type,
                        context.value.clone(),
                    ));
            }
            Ok(())
        }
    }

    let directory = tempdir()?;
    let global_path = directory.path().join("registered-write.xlsx");
    EasyExcel::write::<ConverterRow>(&global_path)
        .register_converter::<String, _>(PrefixConverter::string("first"))
        .register_converter::<String, _>(PrefixConverter::string("latest"))
        .do_write([ConverterRow {
            value: "source".to_owned(),
        }])?;
    let global = EasyExcel::read_sync::<ConverterRow>(&global_path).do_read_sync()?;
    assert_eq!(global[0].value, "latest:source");

    let field_path = directory.path().join("field-precedence.xlsx");
    let observed = Arc::new(Mutex::new(Vec::new()));
    EasyExcel::write::<FieldConverterRow>(&field_path)
        .register_converter::<String, _>(PrefixConverter::string("global"))
        .register_write_handler(OriginalValueProbe(Arc::clone(&observed)))
        .do_write([FieldConverterRow {
            value: "source".to_owned(),
        }])?;
    assert_eq!(
        observed
            .lock()
            .map_err(|_| ExcelError::Format("converter probe poisoned".to_owned()))?
            .as_slice(),
        [(
            Some(CellValue::String("source".to_owned())),
            Some("String"),
            CellValue::String("field:source".to_owned()),
        )]
    );
    let written = EasyExcel::read_sync::<ConverterRow>(&field_path).do_read_sync()?;
    assert_eq!(written[0].value, "field:source");

    let read = EasyExcel::read_sync::<FieldConverterRow>(&global_path)
        .register_converter::<String, _>(PrefixConverter::string("global"))
        .do_read_sync()?;
    assert_eq!(read[0].value, "field:latest:source");
    Ok(())
}

#[test]
fn write_converter_errors_report_physical_sheet_row_column_and_field() -> Result<()> {
    fn row(failing: &str) -> LocatedWriteFailureRow {
        LocatedWriteFailureRow {
            forced: "forced".to_owned(),
            late: "late".to_owned(),
            failing: failing.to_owned(),
        }
    }

    fn assert_location(error: ExcelError, expected_row: u32) {
        match error {
            ExcelError::Data {
                sheet,
                row,
                column,
                field,
                value,
                message,
            } => {
                assert_eq!(sheet, "Diagnostics");
                assert_eq!(row, expected_row);
                assert_eq!(column, Some(0));
                assert_eq!(field, "failing");
                assert!(value.is_empty());
                assert!(message.contains("converter rejected value"));
            }
            other => panic!("expected location-aware conversion error, got {other:?}"),
        }
    }

    let directory = tempdir()?;
    for extension in ["xlsx", "xls", "csv"] {
        let path = directory
            .path()
            .join(format!("located-write-error.{extension}"));
        let error = EasyExcel::write::<LocatedWriteFailureRow>(&path)
            .sheet("Diagnostics")
            .with_bom(false)
            .do_write([row("ok"), row("fail")])
            .expect_err("the second data row must fail conversion");
        assert_location(error, 2);
    }

    let template = directory.path().join("located-write-template.xlsx");
    EasyExcel::write::<Value>(&template)
        .sheet("Diagnostics")
        .need_head(false)
        .do_write([Value("existing".to_owned())])?;
    let template_output = directory.path().join("located-write-template-output.xlsx");
    let template_error = EasyExcel::write::<LocatedWriteFailureRow>(&template_output)
        .with_template(&template)
        .sheet("Diagnostics")
        .need_head(false)
        .do_write([row("ok"), row("fail")])
        .expect_err("template conversion must use the appended physical row");
    assert_location(template_error, 2);

    for extension in ["xlsx", "xls", "csv"] {
        EasyExcel::write::<LocatedWriteFailureRow>(
            directory
                .path()
                .join(format!("excluded-converter.{extension}")),
        )
        .sheet("Diagnostics")
        .with_bom(false)
        .exclude_column_field_names(["failing"])
        .do_write([row("fail")])?;
    }
    Ok(())
}

#[test]
fn sheet_converter_overrides_stateful_workbook_converter() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("stateful-converters.xlsx");
    let mut writer = EasyExcel::write::<ConverterRow>(&path)
        .register_converter::<String, _>(PrefixConverter::string("workbook"))
        .build();
    let workbook_sheet = EasyExcel::writer_sheet::<ConverterRow>("Workbook");
    let override_sheet = EasyExcel::writer_sheet::<ConverterRow>("Override")
        .register_converter::<String, _>(PrefixConverter::string("sheet"));
    writer.write(
        [ConverterRow {
            value: "one".to_owned(),
        }],
        &workbook_sheet,
    )?;
    writer.write(
        [ConverterRow {
            value: "two".to_owned(),
        }],
        &override_sheet,
    )?;
    writer.finish()?;

    let rows = EasyExcel::read_sync::<ConverterRow>(&path)
        .all_sheets()
        .do_read_sync()?;
    assert_eq!(rows[0].value, "workbook:one");
    assert_eq!(rows[1].value, "sheet:two");
    Ok(())
}

/// Java `ExcelWriterBuilder.withTemplate` + `sheet().doWrite` appends onto the template.
#[test]
fn with_template_do_write_appends_and_preserves_other_sheets() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("template.xlsx");
    let output = directory.path().join("template-write.xlsx");

    let mut writer = EasyExcel::write::<Value>(&template).build();
    writer.write(
        [Value("kept".to_owned())],
        &EasyExcel::writer_sheet::<Value>("Sheet1").need_head(false),
    )?;
    writer.write(
        [Value("other".to_owned())],
        &EasyExcel::writer_sheet::<Value>("Sheet2").need_head(false),
    )?;
    writer.finish()?;

    EasyExcel::write::<Value>(&output)
        .with_template(&template)
        .sheet_index(0)
        .need_head(false)
        .do_write([Value("appended".to_owned())])?;

    let sheet1 = EasyExcel::read_sync::<Value>(&output)
        .sheet(0usize)
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(
        sheet1,
        vec![Value("kept".to_owned()), Value("appended".to_owned())]
    );
    let sheet2 = EasyExcel::read_sync::<Value>(&output)
        .sheet(1usize)
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(sheet2, vec![Value("other".to_owned())]);

    let csv = directory.path().join("template-write.csv");
    let error = EasyExcel::write::<Value>(&csv)
        .with_template(&template)
        .do_write([Value("x".to_owned())])
        .expect_err("csv cannot use template");
    assert!(error.to_string().contains("csv cannot use template"));
    Ok(())
}

/// Java `ExcelWriterBuilder.withTemplate(InputStream)` → Rust `with_template_bytes`.
#[test]
fn with_template_bytes_do_write_matches_file_template() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("template-bytes.xlsx");
    let output = directory.path().join("from-bytes.xlsx");

    EasyExcel::write::<Value>(&template)
        .need_head(false)
        .do_write([Value("seed".to_owned())])?;
    let bytes = fs::read(&template)?;

    EasyExcel::write::<Value>(&output)
        .with_template_bytes(bytes)
        .need_head(false)
        .do_write([Value("from-bytes".to_owned())])?;

    let rows = EasyExcel::read_sync::<Value>(&output)
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(
        rows,
        vec![Value("seed".to_owned()), Value("from-bytes".to_owned())]
    );
    Ok(())
}

/// Java `EasyExcel.write(...).withTemplate(...).build()` + `write(data, sheet)` + `finish()`.
#[test]
fn with_template_stateful_writer_appends_on_named_sheet() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("stateful-template.xlsx");
    let output = directory.path().join("stateful-out.xlsx");

    let mut seed = EasyExcel::write::<Value>(&template).build();
    seed.write(
        [Value("alpha".to_owned())],
        &EasyExcel::writer_sheet::<Value>("Data").need_head(false),
    )?;
    seed.write(
        [Value("beta".to_owned())],
        &EasyExcel::writer_sheet::<Value>("Other").need_head(false),
    )?;
    seed.finish()?;

    let mut writer = EasyExcel::write::<Value>(&output)
        .with_template(&template)
        .build();
    writer.write(
        [Value("gamma".to_owned())],
        &EasyExcel::writer_sheet::<Value>("Data").need_head(false),
    )?;
    writer.finish()?;

    let data = EasyExcel::read_sync::<Value>(&output)
        .sheet("Data")
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(
        data,
        vec![Value("alpha".to_owned()), Value("gamma".to_owned())]
    );
    let other = EasyExcel::read_sync::<Value>(&output)
        .sheet("Other")
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(other, vec![Value("beta".to_owned())]);
    Ok(())
}

/// Two-column row used to seed merge + style templates for `with_template` asserts.
#[derive(Debug, Clone, PartialEq, Eq)]
struct PairRow {
    left: String,
    right: String,
}

impl ExcelRow for PairRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[
            ExcelColumn::new("left", "left", Some(0), 0, None),
            ExcelColumn::new("right", "right", Some(1), 0, None),
        ];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        Ok(Self {
            left: row
                .cell(&Self::schema()[0])
                .map_or_else(String::new, CellValue::as_text),
            right: row
                .cell(&Self::schema()[1])
                .map_or_else(String::new, CellValue::as_text),
        })
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        Ok(vec![
            CellValue::String(self.left.clone()),
            CellValue::String(self.right.clone()),
        ])
    }
}

/// Default ZIP path keeps template `styles.xml` / `mergeCells` after `do_write` append.
///
/// Java: `ExcelWriterBuilder.withTemplate` + append — POI keeps styles/merges on the
/// workbook; Rust mirrors that via `TemplatePackage` (not the legacy value-replay seed).
#[test]
fn with_template_do_write_preserves_styles_and_merges() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("styled-template.xlsx");
    let output = directory.path().join("styled-out.xlsx");

    EasyExcel::write::<PairRow>(&template)
        .merge_cells(MergeRange::new(0, 0, 0, 1))
        .head_style(CellStyle::new().bold(true).italic(true))
        .do_write([PairRow {
            left: "seed-l".to_owned(),
            right: "seed-r".to_owned(),
        }])?;

    let styles_before = zip_entry_text(&template, "xl/styles.xml")?;
    let sheet_before = zip_entry_text(&template, "xl/worksheets/sheet1.xml")?;
    assert!(
        styles_before.contains("<b")
            || styles_before.contains("<b/>")
            || styles_before.contains("<b "),
        "template must carry bold style marker: {styles_before}"
    );
    assert!(
        sheet_before.contains("mergeCell") || sheet_before.contains("mergeCells"),
        "template must carry mergeCells: {sheet_before}"
    );

    EasyExcel::write::<Value>(&output)
        .with_template(&template)
        .need_head(false)
        .do_write([Value("appended".to_owned())])?;

    let styles_after = zip_entry_text(&output, "xl/styles.xml")?;
    let sheet_after = zip_entry_text(&output, "xl/worksheets/sheet1.xml")?;
    assert_eq!(
        styles_before, styles_after,
        "ZIP preserve path must leave xl/styles.xml byte-identical"
    );
    assert!(
        sheet_after.contains("mergeCell") || sheet_after.contains("mergeCells"),
        "mergeCells must survive append: {sheet_after}"
    );
    assert!(
        sheet_after.contains("appended"),
        "appended row must be present: {sheet_after}"
    );
    Ok(())
}

/// Creating a sheet absent from the template must not rewrite existing styles/merges.
#[test]
fn with_template_new_sheet_keeps_existing_styles_and_merges() -> Result<()> {
    let directory = tempdir()?;
    let template = directory.path().join("base-template.xlsx");
    let output = directory.path().join("new-sheet-out.xlsx");

    EasyExcel::write::<PairRow>(&template)
        .sheet("Styled")
        .merge_cells(MergeRange::new(0, 0, 0, 1))
        .head_style(CellStyle::new().bold(true))
        .do_write([PairRow {
            left: "a".to_owned(),
            right: "b".to_owned(),
        }])?;

    let styles_before = zip_entry_text(&template, "xl/styles.xml")?;
    let styled_before = zip_entry_text(&template, "xl/worksheets/sheet1.xml")?;

    EasyExcel::write::<Value>(&output)
        .with_template(&template)
        .sheet("Fresh")
        .need_head(false)
        .do_write([Value("on-new".to_owned())])?;

    let styles_after = zip_entry_text(&output, "xl/styles.xml")?;
    let styled_after = zip_entry_text(&output, "xl/worksheets/sheet1.xml")?;
    assert_eq!(
        styles_before, styles_after,
        "styles.xml must stay untouched"
    );
    assert_eq!(
        styled_before, styled_after,
        "existing Styled sheet (incl. mergeCells) must stay byte-identical"
    );

    let fresh = EasyExcel::read_sync::<Value>(&output)
        .sheet("Fresh")
        .head_row_number(0)
        .do_read_sync()?;
    assert_eq!(fresh, vec![Value("on-new".to_owned())]);

    let styled = EasyExcel::read_sync::<PairRow>(&output)
        .sheet("Styled")
        .do_read_sync()?;
    assert_eq!(
        styled,
        vec![PairRow {
            left: "a".to_owned(),
            right: "b".to_owned(),
        }]
    );
    Ok(())
}
