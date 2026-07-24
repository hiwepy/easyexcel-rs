//! Remaining Java `WriteTest` methods not covered in `demo_parity_tests.rs`.
//!
//! Mirrors:
//! - indexWrite / complexHeadWrite / converterWrite
//! - imageWrite / writeCellDataWrite / widthAndHeightWrite
//! - annotationStyleWrite / commentWrite / variableTitleWrite
//! - tableWrite / customHandlerWrite / templateWrite

use chrono::NaiveDate;
use easyexcel::{
    CellStyle, CellValue, ClientAnchorData, CommentData, CoordinateData, EasyExcel, ExcelRow,
    FormulaData, HyperlinkData, HyperlinkType, ImageData, ImageType, RichTextStringData,
    WriteCellData, WriteHandler, WriteWorkbookContext,
};
use tempfile::tempdir;

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn temp_path(name: &str) -> std::path::PathBuf {
    tempdir().unwrap().keep().join(name)
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

/// Java `WriteTest.indexWrite`.
#[test]
fn demo_write_index_write() {
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

/// Java `WriteTest.complexHeadWrite`.
#[test]
fn demo_write_complex_head_write() {
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
    assert!(
        !EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap()
            .is_empty()
    );
}

/// Java `WriteTest.converterWrite`.
#[test]
fn demo_write_converter_write() {
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

/// Java `WriteTest.imageWrite`.
#[test]
fn demo_write_image_write() {
    let img = fixture("converter/img.jpg");
    assert!(
        img.exists(),
        "required Java fixture missing: {}",
        img.display()
    );
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

/// Java `WriteTest.writeCellDataWrite`.
#[test]
fn demo_write_write_cell_data_write() {
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
            CommentData::new()
                .author("Jiaju Zhuang")
                .text("这是一个备注")
                .anchor(
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
    assert!(
        !EasyExcel::read_dynamic_sync(&path)
            .do_read_sync()
            .unwrap()
            .is_empty()
    );
}

/// Java `WriteTest.widthAndHeightWrite`.
#[test]
fn demo_write_width_and_height_write() {
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

/// Java `WriteTest.annotationStyleWrite`.
#[test]
fn demo_write_annotation_style_write() {
    let path = temp_path("annotationStyleWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head_style(CellStyle::default())
        .content_style(CellStyle::default())
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

/// Java `WriteTest.commentWrite` (handler intent: date column carries a note).
#[test]
fn demo_write_comment_write() {
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

/// Java `WriteTest.variableTitleWrite`.
#[test]
fn demo_write_variable_title_write() {
    let path = temp_path("variableTitleWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .head([["字符串标题"], ["日期标题"], ["数字标题"]])
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

/// Java `WriteTest.tableWrite`.
#[test]
fn demo_write_table_write() {
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

/// Java `WriteTest.customHandlerWrite`.
#[test]
fn demo_write_custom_handler_write() {
    #[derive(Default)]
    struct CountingHandler {
        hits: std::sync::Arc<std::sync::atomic::AtomicUsize>,
    }
    impl WriteHandler for CountingHandler {
        fn after_workbook(&mut self, _ctx: &WriteWorkbookContext) -> easyexcel::Result<()> {
            self.hits.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            Ok(())
        }
    }
    let hits = std::sync::Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let path = temp_path("customHandlerWrite.xlsx");
    EasyExcel::write::<WriteDemoData>(&path)
        .register_write_handler(CountingHandler { hits: hits.clone() })
        .sheet("模板")
        .do_write(write_demo_data())
        .unwrap();
    assert!(hits.load(std::sync::atomic::Ordering::Relaxed) >= 1);
}

/// Java `WriteTest.templateWrite` — `withTemplate` + `sheet` + `doWrite`.
///
/// Java:
/// `EasyExcel.write(fileName, DemoData.class).withTemplate(templateFileName).sheet().doWrite(data());`
#[test]
fn demo_write_template_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct DemoData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "日期标题")]
        date: Option<NaiveDate>,
        #[excel(name = "数字标题")]
        double_data: Option<f64>,
    }
    let template = fixture("demo/demo.xlsx");
    assert!(
        template.exists(),
        "required Java fixture missing: {}",
        template.display()
    );
    let template_rows = EasyExcel::read_sync::<DemoData>(&template)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!template_rows.is_empty());
    let path = temp_path("templateWrite.xlsx");
    // Java `.sheet()` selects the first worksheet (sheetNo = 0).
    EasyExcel::write::<WriteDemoData>(&path)
        .with_template(&template)
        .sheet_index(0)
        .do_write(write_demo_data())
        .unwrap();
    // Template Sheet2 must remain (value-level replay of non-target sheets).
    let sheet2 = EasyExcel::read_dynamic_sync(&path)
        .sheet(1usize)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(
        !sheet2.is_empty(),
        "withTemplate must preserve non-target sheets"
    );
    // Java appends a new head + data after the template's last row.
    let all_rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(
        all_rows.len() > template_rows.len() + 1,
        "expected template rows plus appended head/data, got {} (template had {})",
        all_rows.len(),
        template_rows.len()
    );
}
