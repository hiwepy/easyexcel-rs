use std::fs;

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
        .password("read-secret");
    assert_eq!(read.path, PathBuf::from("input.xlsx"));
    assert_eq!(read.options.sheet, SheetSelector::All);
    assert_eq!(read.options.head_row_number, 3);
    assert!(!read.options.ignore_empty_row);
    assert_eq!(read.options.password.as_deref(), Some("read-secret"));

    let sync = EasyExcel::read_sync::<Value>("sync.xlsx")
        .sheet("Values")
        .head_row_number(2)
        .password("sync-secret");
    assert_eq!(sync.path, PathBuf::from("sync.xlsx"));
    assert_eq!(sync.options.sheet, SheetSelector::Name("Values".to_owned()));
    assert_eq!(sync.options.head_row_number, 2);
    assert_eq!(sync.options.password.as_deref(), Some("sync-secret"));

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
