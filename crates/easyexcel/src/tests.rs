use std::collections::BTreeMap;
use std::fs;
use std::io::Cursor;
use std::sync::{Arc, Mutex};

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
    ) -> Result<CellValue> {
        Ok(CellValue::String(format!(
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
    ) -> Result<CellValue> {
        Ok(CellValue::String(format!("field:{}", context.value())))
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct FieldConverterRow {
    #[excel(name = "Value", index = 0, converter = FieldPrefixConverter)]
    value: String,
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
        .loop_merge(LoopMergeStrategy::new(2, 1, 0).expect("loop merge"))
        .head([["Group", "Value"]])
        .password("write-secret")
        .charset("GBK")
        .with_bom(false)
        .register_write_handler(NoopWriteHandler)
        .constant_memory(true);
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
        .do_write([filter_source])?;
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
    EasyExcel::write::<FieldConverterRow>(&field_path)
        .register_converter::<String, _>(PrefixConverter::string("global"))
        .do_write([FieldConverterRow {
            value: "source".to_owned(),
        }])?;
    let written = EasyExcel::read_sync::<ConverterRow>(&field_path).do_read_sync()?;
    assert_eq!(written[0].value, "field:source");

    let read = EasyExcel::read_sync::<FieldConverterRow>(&global_path)
        .register_converter::<String, _>(PrefixConverter::string("global"))
        .do_read_sync()?;
    assert_eq!(read[0].value, "field:latest:source");
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
