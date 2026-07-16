use std::fs;
use std::io::Cursor;

use chrono::NaiveDate;
use tempfile::tempdir;

use super::*;

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

impl WriteHandler for NoopWriteHandler {}

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
fn factories_and_builder_options_match_java_style_chaining() {
    let read = EasyExcel::read::<Value, _>("input.xlsx", Listener::default())
        .sheet(2_usize)
        .all_sheets()
        .head_row_number(3)
        .ignore_empty_row(false)
        .password("read-secret")
        .charset("GBK");
    assert_eq!(read.path, PathBuf::from("input.xlsx"));
    assert_eq!(read.options.sheet, SheetSelector::All);
    assert_eq!(read.options.head_row_number, 3);
    assert!(!read.options.ignore_empty_row);
    assert_eq!(read.options.password.as_deref(), Some("read-secret"));
    assert_eq!(read.options.charset.name(), "GBK");

    let sync = EasyExcel::read_sync::<Value>("sync.xlsx")
        .sheet("Values")
        .head_row_number(2)
        .password("sync-secret")
        .charset(CsvCharset::new("UTF-16BE"));
    assert_eq!(sync.path, PathBuf::from("sync.xlsx"));
    assert_eq!(sync.options.sheet, SheetSelector::Name("Values".to_owned()));
    assert_eq!(sync.options.head_row_number, 2);
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
        .loop_merge(LoopMergeStrategy::new(2, 1, 0).expect("loop merge"))
        .head([["Group", "Value"]])
        .password("write-secret")
        .charset("GBK")
        .with_bom(false)
        .register_write_handler(NoopWriteHandler)
        .constant_memory(true);
    assert_eq!(write.path, PathBuf::from("output.xlsx"));
    assert_eq!(write.options.sheet_name, "Values");
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
    assert_eq!(write.options.password.as_deref(), Some("write-secret"));
    assert_eq!(write.options.charset.name(), "GBK");
    assert!(!write.options.with_bom);
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
    assert!(matches!(
        EasyExcel::write::<Value>("output.xls").do_write(Vec::new()),
        Err(ExcelError::Unsupported(_))
    ));

    let directory = tempdir().expect("temporary directory");
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
