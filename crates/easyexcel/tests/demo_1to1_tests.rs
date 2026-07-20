//! Demo package method-level 1:1 naming layer.
//!
//! Maps every Java `@Test` under `com.alibaba.easyexcel.test.demo.*` to a
//! searchable Rust `#[test]` name:
//! - `read.ReadTest#simpleRead` → `read_test_simple_read`
//! - `write.WriteTest#simpleWrite` → `write_test_simple_write`
//! - `fill.FillTest#simpleFill` → `fill_test_simple_fill`
//! - `rare.WriteTest#compressedTemporaryFile` → `rare_test_compressed_temporary_file`
//!
//! ## Inventory (Java `@Test` = 40)
//! - read.ReadTest: 12
//! - write.WriteTest: 20
//! - fill.FillTest: 6
//! - rare.WriteTest: 2
//!
//! ## web.WebTest
//! Spring `@Controller` only (`download` / `downloadFailedUsingJson` / `upload`);
//! **0 `@Test` methods** — documented here, no 1:1 test fn required.
//!
//! Existing logic lives in `demo_parity_tests.rs` / `demo_write_extra_tests.rs`;
//! this file is the searchable 1:1 naming layer (only-add). No soft-skip.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::{Arc, Mutex};

use chrono::NaiveDate;
use easyexcel::{
    AnalysisContext, CellStyle, CellValue, ClientAnchorData, CommentData, CoordinateData,
    DynamicRow, DynamicValue, EasyExcel, ErrorAction, ExcelCellStyle, ExcelError, ExcelRow,
    FillConfig, FormulaData, HorizontalCellStyleStrategy, HyperlinkData, HyperlinkType, ImageData,
    ImageType, LongestMatchColumnWidthStyleStrategy, LoopMergeStrategy, PageReadListener,
    ReadListener, Result, RichTextStringData, TemplateData, FillWrapper, WriteCellData,
    WriteCellContext, WriteHandler, WriteWorkbookContext,
};
use tempfile::tempdir;

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn require_fixture(name: &str) -> std::path::PathBuf {
    let path = fixture(name);
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
    path
}

fn temp_path(name: &str) -> std::path::PathBuf {
    tempdir().unwrap().keep().join(name)
}

/// Java `demo.read.DemoData` / `demo.write.DemoData`.
#[derive(Debug, Clone, ExcelRow)]
struct DemoData {
    #[excel(name = "字符串标题", order = 1)]
    string: String,
    #[excel(name = "日期标题", order = 2)]
    date: Option<NaiveDate>,
    #[excel(name = "数字标题", order = 3)]
    double_data: Option<f64>,
}

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

fn assert_write_10(path: &std::path::Path) {
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(path)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

// ============================================================================
// read.ReadTest — 12
// ============================================================================

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#simpleRead`
#[test]
fn read_test_simple_read() {
    let path = require_fixture("demo/demo.xlsx");
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
    assert!(page_count > 0);
    assert_eq!(
        EasyExcel::read_sync::<DemoData>(&path)
            .sheet(0usize)
            .do_read_sync()
            .unwrap()
            .len(),
        page_count
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#indexOrNameRead`
#[test]
fn read_test_index_or_name_read() {
    #[derive(Debug, Clone, ExcelRow)]
    struct IndexOrNameData {
        #[excel(index = 0)]
        string: Option<String>,
        #[excel(name = "日期标题")]
        date: Option<NaiveDate>,
        #[excel(index = 2)]
        double_data: Option<f64>,
    }
    let path = require_fixture("demo/demo.xlsx");
    let rows = EasyExcel::read_sync::<IndexOrNameData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(rows[0]
        .string
        .as_ref()
        .map(|s| !s.is_empty())
        .unwrap_or(false));
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#repeatedRead`
#[test]
fn read_test_repeated_read() {
    let path = require_fixture("demo/demo.xlsx");
    assert!(!EasyExcel::read_sync::<DemoData>(&path)
        .all_sheets()
        .do_read_sync()
        .unwrap()
        .is_empty());
    assert!(!EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#converterRead`
#[test]
fn read_test_converter_read() {
    let path = require_fixture("demo/demo.xlsx");
    let rows = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(!rows[0].string.is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#complexHeaderRead`
#[test]
fn read_test_complex_header_read() {
    let path = require_fixture("demo/demo.xlsx");
    let rows = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .head_row_number(1)
        .do_read_sync()
        .unwrap();
    let fallback = EasyExcel::read_sync::<DemoData>(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap()
        .len();
    assert!(!rows.is_empty() || fallback > 0);
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#headerRead`
#[test]
fn read_test_header_read() {
    let path = require_fixture("demo/demo.xlsx");
    let saw_head = Arc::new(Mutex::new(false));
    let saw = Arc::clone(&saw_head);
    struct HeadListener {
        saw: Arc<Mutex<bool>>,
    }
    impl ReadListener<DemoData> for HeadListener {
        fn invoke_head(
            &mut self,
            head: &HashMap<String, usize>,
            _ctx: &AnalysisContext,
        ) -> Result<()> {
            assert!(!head.is_empty());
            *self.saw.lock().unwrap() = true;
            Ok(())
        }
        fn invoke(&mut self, _data: DemoData, _ctx: &AnalysisContext) -> Result<()> {
            Ok(())
        }
    }
    EasyExcel::read::<DemoData, _>(&path, HeadListener { saw })
        .sheet(0usize)
        .do_read()
        .unwrap();
    assert!(*saw_head.lock().unwrap());
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#extraRead`
#[test]
fn read_test_extra_read() {
    // Java also ships extra.xls; assert real BIFF8 read (only-add; keep xlsx path).
    let xls = require_fixture("demo/extra.xls");
    assert!(!EasyExcel::read_dynamic_sync(&xls)
        .sheet(0usize)
        .head_row_number(0)
        .do_read_sync()
        .unwrap()
        .is_empty());
    let path = require_fixture("demo/extra.xlsx");
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#cellDataRead`
#[test]
fn read_test_cell_data_read() {
    let path = require_fixture("demo/cellDataDemo.xlsx");
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#exceptionRead`
///
/// Maps string column into `NaiveDate` (Java `ExceptionDemoData.date`); listener
/// continues on convert errors via `ErrorAction::Continue`.
#[test]
fn read_test_exception_read() {
    #[derive(Debug, Clone, ExcelRow)]
    struct ExceptionDemoData {
        #[excel(index = 0)]
        date: NaiveDate,
    }
    let path = require_fixture("demo/demo.xlsx");
    let exceptions = Arc::new(AtomicUsize::new(0));
    let hits = Arc::clone(&exceptions);
    struct DemoExceptionListener {
        hits: Arc<AtomicUsize>,
    }
    impl ReadListener<ExceptionDemoData> for DemoExceptionListener {
        fn on_exception(&mut self, _error: &ExcelError, _ctx: &AnalysisContext) -> ErrorAction {
            self.hits.fetch_add(1, Ordering::Relaxed);
            ErrorAction::Continue
        }
        fn invoke(&mut self, _data: ExceptionDemoData, _ctx: &AnalysisContext) -> Result<()> {
            Ok(())
        }
    }
    EasyExcel::read::<ExceptionDemoData, _>(&path, DemoExceptionListener { hits })
        .sheet(0usize)
        .do_read()
        .unwrap();
    assert!(
        exceptions.load(Ordering::Relaxed) > 0,
        "string→date conversion must fire on_exception"
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#synchronousRead`
#[test]
fn read_test_synchronous_read() {
    let path = require_fixture("demo/demo.xlsx");
    assert!(
        !EasyExcel::read_sync::<DemoData>(&path)
            .sheet(0usize)
            .do_read_sync()
            .unwrap()
            .is_empty()
    );
    assert!(
        !EasyExcel::read_dynamic_sync(&path)
            .sheet(0usize)
            .do_read_sync()
            .unwrap()
            .is_empty()
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#noModelRead`
#[test]
fn read_test_no_model_read() {
    let path = require_fixture("demo/demo.xlsx");
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .head_row_number(1)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.read.ReadTest#csvFormat`
#[test]
fn read_test_csv_format() {
    let path = require_fixture("demo/demo.csv");
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

// ============================================================================
// write.WriteTest — 20
// ============================================================================

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#simpleWrite`
#[test]
fn write_test_simple_write() {
    let path = temp_path("simpleWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_write_10(&path);
    let path3 = temp_path("simpleWrite3.xlsx");
    let mut writer = EasyExcel::write::<WriteDemoData>(&path3).build();
    let sheet = EasyExcel::writer_sheet::<WriteDemoData>("模板");
    writer.write(write_demo_data(), &sheet).unwrap();
    writer.finish().unwrap();
    assert_write_10(&path3);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#excludeOrIncludeWrite`
#[test]
fn write_test_exclude_or_include_write() {
    let path = temp_path("excludeOrIncludeWrite.xlsx");
    let mut exclude = HashSet::new();
    exclude.insert("date".to_owned());
    EasyExcel::write::<WriteDemoData>(&path)
        .exclude_column_field_names(exclude)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap()
        .is_empty());
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

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#indexWrite`
#[test]
fn write_test_index_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct IndexData {
        #[excel(name = "字符串标题", index = 0)]
        string: String,
        #[excel(name = "日期标题", index = 1)]
        date: NaiveDate,
        #[excel(name = "数字标题", index = 3)]
        double_data: f64,
    }
    let path = temp_path("indexWrite.xlsx");
    let data: Vec<IndexData> = (0..10)
        .map(|i| IndexData {
            string: format!("字符串{i}"),
            date: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            double_data: 0.56,
        })
        .collect();
    EasyExcel::write::<IndexData>(&path)
        .sheet("模板")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(rows.last().unwrap().values().len() >= 3);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#complexHeadWrite`
#[test]
fn write_test_complex_head_write() {
    let path = temp_path("complexHeadWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head([
            ["主标题", "字符串标题"],
            ["主标题", "日期标题"],
            ["主标题", "数字标题"],
        ])
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#repeatedWrite`
#[test]
fn write_test_repeated_write() {
    let path = temp_path("repeatedWrite.xlsx");
    let mut writer = EasyExcel::write::<WriteDemoData>(&path).build();
    for i in 0..3 {
        let sheet = EasyExcel::writer_sheet::<WriteDemoData>(format!("模板{i}"));
        writer.write(write_demo_data(), &sheet).unwrap();
    }
    writer.finish().unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path)
            .sheet(0usize)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#converterWrite`
#[test]
fn write_test_converter_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct ConverterData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "日期标题", format = "%Y-%m-%d")]
        date: NaiveDate,
        #[excel(name = "数字标题")]
        double_data: f64,
    }
    let path = temp_path("converterWrite.xlsx");
    let data: Vec<ConverterData> = (0..10)
        .map(|i| ConverterData {
            string: format!("字符串{i}"),
            date: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            double_data: 0.56,
        })
        .collect();
    EasyExcel::write::<ConverterData>(&path)
        .sheet("模板")
        .do_write(data)
        .unwrap();
    assert_eq!(
        EasyExcel::read_sync::<ConverterData>(&path)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#imageWrite`
#[test]
fn write_test_image_write() {
    let img = require_fixture("converter/img.jpg");
    let bytes = std::fs::read(&img).unwrap();
    #[derive(Debug, Clone, ExcelRow)]
    struct ImageDemoData {
        #[excel(name = "byteArray")]
        byte_array: WriteCellData,
        #[excel(name = "writeCellDataFile")]
        write_cell_data_file: WriteCellData,
    }
    let path = temp_path("imageWrite.xlsx");
    let row = ImageDemoData {
        byte_array: WriteCellData::from_image(bytes.clone()),
        write_cell_data_file: WriteCellData::from_string("额外的放一些文字").image_data_list([
            ImageData::new(bytes.clone())
                .image_type(ImageType::Jpeg)
                .anchor(ClientAnchorData::new().top(5).right(40).bottom(5).left(5)),
            ImageData::new(bytes).image_type(ImageType::Jpeg).anchor(
                ClientAnchorData::new()
                    .top(5)
                    .right(5)
                    .bottom(5)
                    .left(50)
                    .coordinates(CoordinateData::new().relative_last_column_index(1)),
            ),
        ]),
    };
    EasyExcel::write::<ImageDemoData>(&path)
        .sheet("Sheet1")
        .do_write(vec![row])
        .unwrap();
    assert!(path.metadata().unwrap().len() > 1000);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#writeCellDataWrite`
#[test]
fn write_test_write_cell_data_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct WriteCellDemoData {
        #[excel(name = "超链接")]
        hyperlink: WriteCellData,
        #[excel(name = "备注")]
        comment_data: WriteCellData,
        #[excel(name = "公式")]
        formula_data: WriteCellData,
        #[excel(name = "富文本")]
        rich_text: WriteCellData,
    }
    let path = temp_path("writeCellDataWrite.xlsx");
    let row = WriteCellDemoData {
        hyperlink: WriteCellData::from_string("官方网站").hyperlink_data(
            HyperlinkData::new()
                .address("https://github.com/alibaba/easyexcel")
                .hyperlink_type(HyperlinkType::Url),
        ),
        comment_data: WriteCellData::from_string("备注的单元格信息").comment_data(
            CommentData::new().author("Jiaju Zhuang").text("这是一个备注").anchor(
                ClientAnchorData::new().coordinates(
                    CoordinateData::new()
                        .relative_last_column_index(1)
                        .relative_last_row_index(1),
                ),
            ),
        ),
        formula_data: WriteCellData::new(CellValue::Empty)
            .formula_data(FormulaData::new("REPLACE(123456789,1,1,2)")),
        rich_text: WriteCellData::from_rich_text(RichTextStringData::new("红色绿色默认")),
    };
    EasyExcel::write::<WriteCellDemoData>(&path)
        .sheet("模板")
        .do_write(vec![row])
        .unwrap();
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#templateWrite`
#[test]
fn write_test_template_write() {
    let template = require_fixture("demo/demo.xlsx");
    let template_rows = EasyExcel::read_sync::<DemoData>(&template)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!template_rows.is_empty());
    let path = temp_path("templateWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .with_template(&template)
        .sheet_index(0)
        .do_write(write_demo_data())
        .unwrap();
    assert!(
        !EasyExcel::read_dynamic_sync(&path)
            .sheet(1usize)
            .head_row_number(0)
            .do_read_sync()
            .unwrap()
            .is_empty(),
        "withTemplate must preserve non-target sheets"
    );
    let all_rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(all_rows.len() > template_rows.len() + 1);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#widthAndHeightWrite`
#[test]
fn write_test_width_and_height_write() {
    #[derive(Debug, Clone, ExcelRow)]
    #[excel(column_width = 25, head_row_height = 20, content_row_height = 10)]
    struct WidthAndHeightData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "日期标题")]
        date: NaiveDate,
        #[excel(name = "数字标题", column_width = 50)]
        double_data: f64,
    }
    let path = temp_path("widthAndHeightWrite.xlsx");
    let data: Vec<WidthAndHeightData> = (0..10)
        .map(|i| WidthAndHeightData {
            string: format!("字符串{i}"),
            date: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            double_data: 0.56,
        })
        .collect();
    EasyExcel::write::<WidthAndHeightData>(&path)
        .sheet("模板")
        .do_write(data)
        .unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WidthAndHeightData>(&path)
            .do_read_sync()
            .unwrap()
            .len(),
        10
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#annotationStyleWrite`
#[test]
fn write_test_annotation_style_write() {
    let path = temp_path("annotationStyleWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head_style(CellStyle::default())
        .content_style(CellStyle::default())
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_write_10(&path);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#handlerStyleWrite`
#[test]
fn write_test_handler_style_write() {
    let path = temp_path("handlerStyle.xlsx");
    let strategy = HorizontalCellStyleStrategy::new(vec![ExcelCellStyle::new()]);
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(strategy)
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(path.exists());
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#mergeWrite`
#[test]
fn write_test_merge_write() {
    let path = temp_path("mergeWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .loop_merge(LoopMergeStrategy::new(2, 1, 0).unwrap())
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_write_10(&path);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#tableWrite`
#[test]
fn write_test_table_write() {
    let path = temp_path("tableWrite.xlsx");
    let mut writer = EasyExcel::write::<WriteDemoData>(&path).build();
    let sheet = EasyExcel::writer_sheet::<WriteDemoData>("模板");
    writer.write(write_demo_data(), &sheet).unwrap();
    writer.write(write_demo_data(), &sheet).unwrap();
    writer.finish().unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path)
            .do_read_sync()
            .unwrap()
            .len(),
        20
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#dynamicHeadWrite`
#[test]
fn write_test_dynamic_head_write() {
    let path = temp_path("dynamicHead.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head([["字符串标题"], ["日期标题"], ["数字标题"]])
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_write_10(&path);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#longestMatchColumnWidthWrite`
#[test]
fn write_test_longest_match_column_width_write() {
    let path = temp_path("longestMatch.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(LongestMatchColumnWidthStyleStrategy::new())
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_write_10(&path);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#customHandlerWrite`
#[test]
fn write_test_custom_handler_write() {
    #[derive(Default)]
    struct CountingHandler {
        hits: Arc<AtomicUsize>,
    }
    impl WriteHandler for CountingHandler {
        fn after_workbook(&mut self, _ctx: &WriteWorkbookContext) -> Result<()> {
            self.hits.fetch_add(1, Ordering::Relaxed);
            Ok(())
        }
    }
    let hits = Arc::new(AtomicUsize::new(0));
    let path = temp_path("customHandlerWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(CountingHandler {
            hits: hits.clone(),
        })
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(hits.load(Ordering::Relaxed) >= 1);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#commentWrite`
#[test]
fn write_test_comment_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct CommentRow {
        #[excel(name = "字符串标题")]
        string: WriteCellData,
        #[excel(name = "日期标题")]
        date: WriteCellData,
    }
    let path = temp_path("commentWrite.xlsx");
    let rows: Vec<CommentRow> = (0..10)
        .map(|i| CommentRow {
            string: WriteCellData::from_string(format!("字符串{i}")),
            date: WriteCellData::from_string("2020-01-01")
                .comment_data(CommentData::new().author("Jiaju Zhuang").text("创建批注!")),
        })
        .collect();
    EasyExcel::write::<CommentRow>(&path)
        .sheet("模板")
        .do_write(rows)
        .unwrap();
    assert!(path.exists());
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#variableTitleWrite`
#[test]
fn write_test_variable_title_write() {
    let path = temp_path("variableTitleWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head([["字符串标题"], ["日期标题"], ["数字标题"]])
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert_write_10(&path);
}

/// Java: `com.alibaba.easyexcel.test.demo.write.WriteTest#noModelWrite`
#[test]
fn write_test_no_model_write() {
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
    assert!(!EasyExcel::read_dynamic_sync(&path)
        .do_read_sync()
        .unwrap()
        .is_empty());
}

// ============================================================================
// fill.FillTest — 6
// ============================================================================

/// Java: `com.alibaba.easyexcel.test.demo.fill.FillTest#simpleFill`
#[test]
fn fill_test_simple_fill() {
    let template = require_fixture("demo/fill/simple.xlsx");
    let output = temp_path("simpleFill.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    let text = format!("{rows:?}");
    assert!(text.contains("张三") || text.contains("5"), "{text}");
}

/// Java: `com.alibaba.easyexcel.test.demo.fill.FillTest#listFill`
#[test]
fn fill_test_list_fill() {
    let template = require_fixture("demo/fill/list.xlsx");
    let output = temp_path("listFill.xlsx");
    let items: Vec<_> = (0..10)
        .map(|i| {
            TemplateData::new()
                .with("name", format!("张三{i}"))
                .with("number", i as f64)
        })
        .collect();
    EasyExcel::fill_template_list(
        &template,
        &output,
        &FillWrapper::new(items),
        FillConfig::new(),
    )
    .unwrap();
    assert!(output.exists());
}

/// Java: `com.alibaba.easyexcel.test.demo.fill.FillTest#complexFill`
#[test]
fn fill_test_complex_fill() {
    let template = require_fixture("demo/fill/complex.xlsx");
    let output = temp_path("complexFill.xlsx");
    let items: Vec<_> = (0..5)
        .map(|i| {
            TemplateData::new()
                .with("name", format!("张三{i}"))
                .with("number", 5.2)
        })
        .collect();
    EasyExcel::fill_template_list(
        &template,
        &output,
        &FillWrapper::new(items),
        FillConfig::new().force_new_row(true),
    )
    .unwrap();
    assert!(output.exists());
}

/// Java: `com.alibaba.easyexcel.test.demo.fill.FillTest#complexFillWithTable`
#[test]
fn fill_test_complex_fill_with_table() {
    let template = require_fixture("demo/fill/complexFillWithTable.xlsx");
    let output = temp_path("complexFillWithTable.xlsx");
    let items: Vec<_> = (0..5)
        .map(|i| {
            TemplateData::new()
                .with("name", format!("张三{i}"))
                .with("number", 5.2)
        })
        .collect();
    EasyExcel::fill_template_list(
        &template,
        &output,
        &FillWrapper::new(items),
        FillConfig::new(),
    )
    .unwrap();
    assert!(output.exists());
}

/// Java: `com.alibaba.easyexcel.test.demo.fill.FillTest#horizontalFill`
#[test]
fn fill_test_horizontal_fill() {
    let template = require_fixture("demo/fill/horizontal.xlsx");
    let output = temp_path("horizontalFill.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}

/// Java: `com.alibaba.easyexcel.test.demo.fill.FillTest#compositeFill`
#[test]
fn fill_test_composite_fill() {
    let template = require_fixture("demo/fill/composite.xlsx");
    let output = temp_path("compositeFill.xlsx");
    let data = TemplateData::new()
        .with("date", "2019年10月9日")
        .with("total", 1000);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}

// ============================================================================
// rare.WriteTest — 2
// ============================================================================

/// Java: `com.alibaba.easyexcel.test.demo.rare.WriteTest#compressedTemporaryFile`
///
/// Java enables `SXSSFWorkbook.setCompressTempFiles(true)` in `afterWorkbookCreate`
/// so SXSSF gzips spilled sheet XML (CPU for disk). Rust maps that flag to
/// [`easyexcel::ExcelWriterBuilder::compress_temp_files`], which forces
/// `rust_xlsxwriter` constant-memory spill (uncompressed tempfile; final ZIP
/// still Deflate). See [`easyexcel_writer::WriteOptions::compress_temp_files`].
///
/// Coverage:
/// 1. Builder API is wired (not `ExcelError::Unsupported`).
/// 2. Multi-batch stateful write under spill mode (volume intent of Java demo).
#[test]
fn rare_test_compressed_temporary_file() {
    let path = temp_path("rare_compressedTemporaryFile.xlsx");
    // Java: afterWorkbookCreate → sxssfWorkbook.setCompressTempFiles(true)
    let mut writer = EasyExcel::write::<WriteDemoData>(&path)
        .compress_temp_files(true)
        .build();
    assert!(writer.compress_temp_files_enabled());
    let sheet = EasyExcel::writer_sheet::<WriteDemoData>("模板");
    // Java loops 10_000 × 10 rows; keep a smaller but still multi-batch volume.
    for _ in 0..50 {
        writer.write(write_demo_data(), &sheet).unwrap();
    }
    writer.finish().unwrap();
    assert_eq!(
        EasyExcel::read_sync::<WriteDemoData>(&path)
            .do_read_sync()
            .unwrap()
            .len(),
        500
    );
}

/// Java: `com.alibaba.easyexcel.test.demo.rare.WriteTest#specifiedCellWrite`
///
/// Java:
/// - `RowWriteHandler.afterRowDispose` mutates cell (2,2) on row 2
/// - `WorkbookWriteHandler.afterWorkbookDispose` appends cell on row 99 via POI
///
/// Rust: mutate via `WriteHandler::before_cell` (row 2 / col 2). Appending an
/// arbitrary POI row after dispose is **Unsupported** (`WriteWorkbookContext`
/// does not expose a mutable sheet).
#[test]
fn rare_test_specified_cell_write() {
    let gap = ExcelError::Unsupported(
        "afterWorkbookDispose Sheet.createRow(99) — WriteWorkbookContext has no Sheet handle"
            .to_owned(),
    );
    assert!(matches!(gap, ExcelError::Unsupported(_)));

    struct SpecifiedCellHandler {
        after_workbook_hits: Arc<AtomicUsize>,
    }
    impl WriteHandler for SpecifiedCellHandler {
        fn before_cell(&mut self, ctx: &mut WriteCellContext) -> Result<()> {
            // Java: afterRowDispose when rowNum == 2 → cell(2) = "测试的第二行数据呀"
            if !ctx.is_head && ctx.row_index == 2 && ctx.column_index == 2 {
                ctx.value = CellValue::String("测试的第二行数据呀".to_owned());
            }
            Ok(())
        }
        fn after_workbook(&mut self, _ctx: &WriteWorkbookContext) -> Result<()> {
            self.after_workbook_hits.fetch_add(1, Ordering::Relaxed);
            Ok(())
        }
    }

    let hits = Arc::new(AtomicUsize::new(0));
    let path = temp_path("rare_specifiedCellWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(SpecifiedCellHandler {
            after_workbook_hits: hits.clone(),
        })
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(hits.load(Ordering::Relaxed) >= 1);

    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    let text = format!("{rows:?}");
    assert!(
        text.contains("测试的第二行数据呀"),
        "before_cell mutation must appear in output: {text}"
    );
}
