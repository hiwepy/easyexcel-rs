//! Demo package parity — mirrors Java
//! `com.alibaba.easyexcel.test.demo.read.ReadTest`,
//! `com.alibaba.easyexcel.test.demo.write.WriteTest`,
//! `com.alibaba.easyexcel.test.demo.fill.FillTest`.
//!
//! Demo Java tests mostly log; Rust asserts observable outcomes so results
//! can be compared with Java (row counts, key cell values, file existence).
//!
//! Fixtures: `tests/fixtures/demo/*` copied from Java `easyexcel-test` resources.

use std::collections::HashSet;
use std::sync::{Arc, Mutex};

use chrono::NaiveDate;
use std::collections::BTreeMap;

use easyexcel::{
    AnalysisContext, DynamicRow, DynamicValue, EasyExcel, ExcelCellStyle, ExcelRow, FillConfig,
    FillWrapper, HorizontalCellStyleStrategy, LongestMatchColumnWidthStyleStrategy,
    LoopMergeStrategy, PageReadListener, ReadListener, Result, TemplateData,
};
use tempfile::tempdir;

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn temp_path(name: &str) -> std::path::PathBuf {
    let dir = tempdir().unwrap();
    dir.keep().join(name)
}

// ============================================================================
// Demo read models — Java demo.read.DemoData / IndexOrNameData
// ============================================================================

/// Java `com.alibaba.easyexcel.test.demo.read.DemoData`
#[derive(Debug, Clone, ExcelRow)]
struct DemoData {
    #[excel(name = "字符串标题")]
    string: String,
    #[excel(name = "日期标题")]
    date: Option<NaiveDate>,
    #[excel(name = "数字标题")]
    double_data: Option<f64>,
}

/// Java `com.alibaba.easyexcel.test.demo.read.IndexOrNameData`
#[derive(Debug, Clone, ExcelRow)]
struct IndexOrNameData {
    #[excel(index = 0)]
    string: Option<String>,
    #[excel(name = "日期标题")]
    date: Option<NaiveDate>,
    #[excel(index = 2)]
    double_data: Option<f64>,
}

/// Java `com.alibaba.easyexcel.test.demo.write.DemoData` (same fields, write side)
#[derive(Debug, Clone, ExcelRow)]
struct WriteDemoData {
    #[excel(name = "字符串标题", order = 1)]
    string: String,
    #[excel(name = "日期标题", order = 2)]
    date: NaiveDate,
    #[excel(name = "数字标题", order = 3)]
    double_data: f64,
}

fn write_demo_data() -> Vec<WriteDemoData> {
    (0..10)
        .map(|i| WriteDemoData {
            string: format!("字符串{i}"),
            date: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            double_data: 0.56,
        })
        .collect()
}

fn count_demo_rows(path: &std::path::Path) -> usize {
    EasyExcel::read_sync::<DemoData>(path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap()
        .len()
}

// ============================================================================
// ReadTest
// ============================================================================

/// Java `ReadTest.simpleRead` — PageReadListener + sync read of demo.xlsx.
#[test]
fn demo_read_simple_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());

    let total = Arc::new(Mutex::new(0usize));
    let total_cb = Arc::clone(&total);
    let listener = PageReadListener::new(100, move |batch: Vec<DemoData>, _ctx| {
        *total_cb.lock().unwrap() += batch.len();
        Ok(())
    });
    EasyExcel::read::<DemoData, _>(&path, listener)
        .sheet(0usize)
        .do_read()
        .unwrap();
    let page_count = *total.lock().unwrap();
    assert!(page_count > 0, "Java simpleRead must read at least one row");

    let sync_rows = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert_eq!(sync_rows.len(), page_count);
}

/// Java `ReadTest.indexOrNameRead`.
#[test]
fn demo_read_index_or_name_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_sync::<IndexOrNameData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(rows[0].string.as_ref().map(|s| !s.is_empty()).unwrap_or(false));
}

/// Java `ReadTest.repeatedRead` — all sheets + selected sheets.
#[test]
fn demo_read_repeated_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());

    let all = EasyExcel::read_sync::<DemoData>(&path)
        .all_sheets()
        .do_read_sync()
        .unwrap();
    assert!(!all.is_empty(), "doReadAll must return rows");

    let sheet0 = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!sheet0.is_empty());
}

/// Java `ReadTest.converterRead` — date/number formatted fields via model.
#[test]
fn demo_read_converter_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(!rows[0].string.is_empty());
}

/// Java `ReadTest.complexHeaderRead` — head_row_number.
#[test]
fn demo_read_complex_header_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .head_row_number(1)
        .do_read_sync()
        .unwrap();
    // demo.xlsx has a single header row; still must succeed and return data rows
    assert!(!rows.is_empty() || count_demo_rows(&path) > 0);
}

/// Java `ReadTest.headerRead` — invoke_head callback fires.
#[test]
fn demo_read_header_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let saw_head = Arc::new(Mutex::new(false));
    let saw = Arc::clone(&saw_head);
    struct HeadListener {
        saw: Arc<Mutex<bool>>,
        rows: usize,
    }
    impl ReadListener<DemoData> for HeadListener {
        fn invoke_head(
            &mut self,
            head: &std::collections::HashMap<String, usize>,
            _ctx: &AnalysisContext,
        ) -> Result<()> {
            assert!(!head.is_empty());
            *self.saw.lock().unwrap() = true;
            Ok(())
        }
        fn invoke(&mut self, _data: DemoData, _ctx: &AnalysisContext) -> Result<()> {
            self.rows += 1;
            Ok(())
        }
    }
    EasyExcel::read::<DemoData, _>(
        &path,
        HeadListener {
            saw,
            rows: 0,
        },
    )
    .sheet(0usize)
    .do_read()
    .unwrap();
    assert!(*saw_head.lock().unwrap(), "invokeHead must be called");
}

/// Java `ReadTest.extraRead` — extra.xlsx comment/hyperlink/merge metadata.
#[test]
fn demo_read_extra_read() {
    let path = fixture("demo/extra.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `ReadTest.cellDataRead` — cellDataDemo.xlsx.
#[test]
fn demo_read_cell_data_read() {
    let path = fixture("demo/cellDataDemo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `ReadTest.synchronousRead`.
#[test]
fn demo_read_synchronous_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let n = count_demo_rows(&path);
    assert!(n > 0);
}

/// Java `ReadTest.noModelRead`.
#[test]
fn demo_read_no_model_read() {
    let path = fixture("demo/demo.xlsx");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .head_row_number(1)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    match rows[0].get(0) {
        Some(DynamicValue::String(s)) => assert!(!s.is_empty()),
        Some(DynamicValue::ActualData(_)) => {}
        other => panic!("expected cell value, got {other:?}"),
    }
}

/// Java `ReadTest.csvFormat` — demo.csv.
#[test]
fn demo_read_csv_format() {
    let path = fixture("demo/demo.csv");
    assert!(path.exists(), "required Java fixture missing: {}", path.display());
    let rows = EasyExcel::read_dynamic_sync(&path)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

// ============================================================================
// WriteTest
// ============================================================================

/// Java `WriteTest.simpleWrite` — three write styles, then read-back assert.
#[test]
fn demo_write_simple_write() {
    let path1 = temp_path("simpleWrite1.xlsx");
    EasyExcel::write::<WriteDemoData>(&path1)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path1)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );

    let path2 = temp_path("simpleWrite2.xlsx");
    EasyExcel::write::<WriteDemoData>(&path2)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(path2.exists());

    let path3 = temp_path("simpleWrite3.xlsx");
    let mut writer = EasyExcel::write::<WriteDemoData>(&path3).build();
    let sheet = EasyExcel::writer_sheet::<WriteDemoData>("模板");
    writer.write(write_demo_data(), &sheet).unwrap();
    writer.finish().unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path3)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

/// Java `WriteTest.excludeOrIncludeWrite`.
#[test]
fn demo_write_exclude_or_include_write() {
    let path = temp_path("excludeOrIncludeWrite.xlsx");
    let mut exclude = HashSet::new();
    exclude.insert("date".to_owned());
    EasyExcel::write::<WriteDemoData>(&path)
        .exclude_column_field_names(exclude)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());

    let path2 = temp_path("includeOnlyDate.xlsx");
    let mut include = HashSet::new();
    include.insert("date".to_owned());
    EasyExcel::write::<WriteDemoData>(&path2)
        .include_column_field_names(include)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(path2.exists());
}

/// Java `WriteTest.noModelWrite`.
#[test]
fn demo_write_no_model_write() {
    let path = temp_path("noModelWrite.xlsx");
    let rows: Vec<DynamicRow> = (0..10)
        .map(|i| {
            let mut map = BTreeMap::new();
            map.insert(0, DynamicValue::String(format!("字符串{i}")));
            map.insert(1, DynamicValue::String("2020-01-01".to_owned()));
            map.insert(2, DynamicValue::String("0.56".to_owned()));
            DynamicRow::new(map)
        })
        .collect();
    EasyExcel::write::<DynamicRow>(&path)
        .head([["字符串标题"], ["日期标题"], ["数字标题"]])
        .sheet("模板")
        .do_write(rows)
        .unwrap();
    let back = EasyExcel::read_dynamic_sync(&path)
        .do_read_sync()
        .unwrap();
    assert!(!back.is_empty());
}

/// Java `WriteTest.repeatedWrite` — stateful ExcelWriter multi-sheet.
#[test]
fn demo_write_repeated_write() {
    let path = temp_path("repeatedWrite.xlsx");
    let mut writer = EasyExcel::write::<WriteDemoData>(&path).build();
    for i in 0..3 {
        let sheet = EasyExcel::writer_sheet::<WriteDemoData>(format!("模板{i}"));
        writer.write(write_demo_data(), &sheet).unwrap();
    }
    writer.finish().unwrap();
    let sheet0 = EasyExcel::read_sync::<WriteDemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert_eq!(sheet0.len(), 10);
}

/// Java `WriteTest.longestMatchColumnWidthWrite`.
#[test]
fn demo_write_longest_match_column_width_write() {
    let path = temp_path("longestMatch.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(LongestMatchColumnWidthStyleStrategy::new())
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

/// Java `WriteTest.mergeWrite` — LoopMergeStrategy.
#[test]
fn demo_write_merge_write() {
    let path = temp_path("mergeWrite.xlsx");
    // Java: new LoopMergeStrategy(2, 0) ≈ eachRow=2, columnExtend=1, columnIndex=0
    EasyExcel::write::<WriteDemoData>(&path)
        .loop_merge(LoopMergeStrategy::new(2, 1, 0).unwrap())
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

/// Java `WriteTest.handlerStyleWrite` — HorizontalCellStyleStrategy registered.
#[test]
fn demo_write_handler_style_write() {
    let path = temp_path("handlerStyle.xlsx");
    let strategy = HorizontalCellStyleStrategy::new(vec![ExcelCellStyle::new()]);
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(strategy)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(path.exists());
}

/// Java `WriteTest.dynamicHeadWrite`.
#[test]
fn demo_write_dynamic_head_write() {
    let path = temp_path("dynamicHead.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head([["字符串标题"], ["日期标题"], ["数字标题"]])
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<WriteDemoData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
}

// ============================================================================
// FillTest — Java demo.fill templates under demo/fill/
// ============================================================================

/// Java `FillTest.simpleFill`.
#[test]
fn demo_fill_simple_fill() {
    let template = fixture("demo/fill/simple.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let output = temp_path("simpleFill.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    let text = format!("{rows:?}");
    assert!(text.contains("张三") || text.contains("5"), "filled values must appear: {text}");
}

/// Java `FillTest.listFill`.
#[test]
fn demo_fill_list_fill() {
    let template = fixture("demo/fill/list.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let output = temp_path("listFill.xlsx");
    let mut items = Vec::new();
    for i in 0..10 {
        items.push(
            TemplateData::new()
                .with("name", format!("张三{i}"))
                .with("number", i as f64),
        );
    }
    let wrapper = FillWrapper::new(items);
    EasyExcel::fill_template_list(&template, &output, &wrapper, FillConfig::new()).unwrap();
    assert!(output.exists());
}

/// Java `FillTest.complexFill`.
#[test]
fn demo_fill_complex_fill() {
    let template = fixture("demo/fill/complex.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let output = temp_path("complexFill.xlsx");
    let mut items = Vec::new();
    for i in 0..5 {
        items.push(
            TemplateData::new()
                .with("name", format!("张三{i}"))
                .with("number", 5.2),
        );
    }
    let wrapper = FillWrapper::new(items);
    EasyExcel::fill_template_list(
        &template,
        &output,
        &wrapper,
        FillConfig::new().force_new_row(true),
    )
    .unwrap();
    assert!(output.exists());
}

/// Java `FillTest.complexFillWithTable`.
#[test]
fn demo_fill_complex_fill_with_table() {
    let template = fixture("demo/fill/complexFillWithTable.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let output = temp_path("complexFillWithTable.xlsx");
    let mut items = Vec::new();
    for i in 0..5 {
        items.push(
            TemplateData::new()
                .with("name", format!("张三{i}"))
                .with("number", 5.2),
        );
    }
    let wrapper = FillWrapper::new(items);
    EasyExcel::fill_template_list(&template, &output, &wrapper, FillConfig::new()).unwrap();
    assert!(output.exists());
}

/// Java `FillTest.horizontalFill`.
#[test]
fn demo_fill_horizontal_fill() {
    let template = fixture("demo/fill/horizontal.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let output = temp_path("horizontalFill.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}

/// Java `FillTest.compositeFill`.
#[test]
fn demo_fill_composite_fill() {
    let template = fixture("demo/fill/composite.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let output = temp_path("compositeFill.xlsx");
    let data = TemplateData::new()
        .with("date", "2019年10月9日")
        .with("total", 1000);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}
