//! Temp-package *contract* tests — EasyExcel API only (not pure POI / hardcoded paths).
//!
//! Selected from Java `com.alibaba.easyexcel.test.temp.*` where the scenario is
//! portable. Skips machine-local paths (`D:\\test\\...`) and raw POI suites.
//!
//! Fill-oriented contracts live in `temp_fill_contract_tests.rs`.

use easyexcel::{DynamicRow, EasyExcel, ExcelCellStyle, ExcelRow, HorizontalCellStyleStrategy};
use tempfile::tempdir;

fn temp_path(name: &str) -> std::path::PathBuf {
    tempdir().unwrap().keep().join(name)
}

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn assert_fixture(path: &std::path::Path) {
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
}

/// Java `temp.csv.CsvReadTest.csvWrite` / `writev2` — CSV write then read.
#[test]
fn temp_csv_write_and_read() {
    #[derive(Debug, Clone, ExcelRow)]
    struct CsvData {
        #[excel(name = "userId")]
        user_id: String,
        #[excel(name = "userName")]
        user_name: String,
    }
    let path = temp_path("csvWrite1.csv");
    let data: Vec<CsvData> = (0..10)
        .map(|i| CsvData {
            user_id: format!("userId{i}"),
            user_name: format!("userName{i}"),
        })
        .collect();
    EasyExcel::write::<CsvData>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_sync::<CsvData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].user_id, "userId0");
    assert_eq!(rows[0].user_name, "userName0");

    let dynamic = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!dynamic.is_empty());
}

/// Java `temp.csv.CsvReadTest.writev2` / `data()` — CsvData model with comma + ignore.
#[test]
fn temp_csv_csvdata_model_round_trip() {
    #[derive(Debug, Clone, ExcelRow)]
    struct CsvData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "数字标题")]
        double_data: f64,
        #[excel(ignore)]
        ignore: String,
    }
    let path = temp_path("csv_csvdata.csv");
    let data: Vec<CsvData> = (0..10)
        .map(|i| CsvData {
            string: format!("字符,串{i}"),
            double_data: 0.56,
            ignore: format!("忽略{i}"),
        })
        .collect();
    EasyExcel::write::<CsvData>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_sync::<CsvData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].string, "字符,串0");
    assert!((rows[0].double_data - 0.56).abs() < 1e-9);
    assert!(rows[0].ignore.is_empty() || rows[0].ignore == String::default());
}

/// Java `temp.csv.CsvReadTest` intent — BOM / no-BOM CSV fixtures readable.
#[test]
fn temp_csv_bom_fixtures_readable() {
    for name in ["bom/office_bom.csv", "bom/no_bom.csv", "demo/demo.csv"] {
        let path = fixture(name);
        assert_fixture(&path);
        let rows = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
        assert!(!rows.is_empty(), "{name} must yield rows");
    }
}

/// Java `temp.read.HeadReadTest` intent — head row readable from demo.xlsx.
#[test]
fn temp_head_read_demo_xlsx() {
    let path = fixture("demo/demo.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.read.HeadReadTest` — `ignoreEmptyRow(false)` still returns data.
#[test]
fn temp_head_read_ignore_empty_row_false() {
    let path = fixture("demo/demo.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .ignore_empty_row(false)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.Xls03Test` intent — .xls readable via calamine path.
#[test]
fn temp_xls03_read_fixture() {
    let path = fixture("xls/converter03.xls");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.Xls03Test` / compatibility — multiplesheets.xls readable.
#[test]
fn temp_xls03_multiplesheets_readable() {
    let path = fixture("xls/multiplesheets.xls");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .all_sheets()
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.large.TempLargeDataTest` / WriteLarge intent — 5000-row round trip.
#[test]
fn temp_large_write_read_xlsx() {
    #[derive(Debug, Clone, ExcelRow)]
    struct LargeRow {
        #[excel(name = "id")]
        id: i64,
        #[excel(name = "name")]
        name: String,
    }
    let path = temp_path("temp_large.xlsx");
    let data: Vec<LargeRow> = (0..5_000)
        .map(|i| LargeRow {
            id: i,
            name: format!("name{i}"),
        })
        .collect();
    EasyExcel::write::<LargeRow>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_sync::<LargeRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 5_000);
    assert_eq!(rows[4999].name, "name4999");
}

/// Java `temp.WriteLargeTest.test2` — batched stateful writes then read back.
#[test]
fn temp_write_large_batched_xlsx() {
    #[derive(Debug, Clone, ExcelRow)]
    struct LargeRow {
        #[excel(name = "c0")]
        c0: String,
        #[excel(name = "c1")]
        c1: String,
    }
    let path = temp_path("temp_write_large_batched.xlsx");
    let mut writer = EasyExcel::write::<LargeRow>(&path).build();
    let sheet = EasyExcel::writer_sheet::<LargeRow>("Sheet1");
    for batch in 0..10 {
        let rows: Vec<LargeRow> = (0..100)
            .map(|i| LargeRow {
                c0: format!("batch{batch}-row{i}"),
                c1: format!("这是测试字段{i}"),
            })
            .collect();
        writer.write(rows, &sheet).unwrap();
    }
    writer.finish().unwrap();
    let rows = EasyExcel::read_sync::<LargeRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1_000);
    assert_eq!(rows[0].c0, "batch0-row0");
    assert_eq!(rows[999].c0, "batch9-row99");
}

/// Java `temp.large.TempLargeDataTest` — large07 fixture present (74MB; do not sync-read all).
///
/// Full-file sync read is intentionally avoided; assertable contract is fixture
/// availability + OOXML package magic. Batch write/read covered by
/// `temp_write_large_batched_xlsx` / `temp_large_write_read_xlsx`.
#[test]
fn temp_large_fixture_present() {
    let path = fixture("large/large07.xlsx");
    assert_fixture(&path);
    let mut header = [0u8; 4];
    let mut file = std::fs::File::open(&path).unwrap();
    use std::io::Read;
    file.read_exact(&mut header).unwrap();
    assert_eq!(&header, b"PK\x03\x04", "large07 must be OOXML zip");
    assert!(
        path.metadata().unwrap().len() > 1_000_000,
        "large07 fixture should be multi-MB"
    );
}

/// Java `temp.StyleTest` intent — style registration does not break write/read.
#[test]
fn temp_style_write_read() {
    #[derive(Debug, Clone, ExcelRow)]
    #[excel(column_width = 20)]
    struct StyleRow {
        #[excel(name = "col")]
        col: String,
    }
    let path = temp_path("temp_style.xlsx");
    EasyExcel::write::<StyleRow>(&path)
        .sheet("Sheet1")
        .do_write(vec![StyleRow {
            col: "styled".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<StyleRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows[0].col, "styled");
}

/// Java `temp.WriteV34Test` / `Lock2Test.write` — HorizontalCellStyleStrategy write.
#[test]
fn temp_style_horizontal_strategy_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct DemoData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "数字标题")]
        double_data: f64,
    }
    let path = temp_path("temp_style_handler.xlsx");
    let strategy = HorizontalCellStyleStrategy::new(vec![ExcelCellStyle::new()]);
    EasyExcel::write::<DemoData>(&path)
        .register_write_handler(strategy)
        .sheet("模板")
        .do_write(vec![
            DemoData {
                string: "字符串0".into(),
                double_data: 0.56,
            };
            10
        ])
        .unwrap();
    let rows = EasyExcel::read_sync::<DemoData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].string, "字符串0");
}

/// Java `temp.fill.FillTempTest` intent — simple template fill still works.
#[test]
fn temp_fill_simple_template() {
    let template = fixture("demo/fill/simple.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill.xlsx");
    let data = easyexcel::TemplateData::new()
        .with("name", "李四")
        .with("number", 3.14);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}

/// Java `temp.simple.RepeatTest` intent — repeated sheet writes.
#[test]
fn temp_repeat_multi_sheet_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Row {
        #[excel(name = "v")]
        v: i32,
    }
    let path = temp_path("temp_repeat.xlsx");
    let mut writer = EasyExcel::write::<Row>(&path).build();
    for i in 0..3 {
        let sheet = EasyExcel::writer_sheet::<Row>(format!("s{i}"));
        writer
            .write(vec![Row { v: i }, Row { v: i + 10 }], &sheet)
            .unwrap();
    }
    writer.finish().unwrap();
    let all = EasyExcel::read_sync::<Row>(&path)
        .all_sheets()
        .do_read_sync()
        .unwrap();
    assert_eq!(all.len(), 6);
}

/// Java `temp.simple.RepeatTest` intent — multi-sheet fixture readable by index.
#[test]
fn temp_repeat_multiplesheets_fixture() {
    let path = fixture("multiplesheets/multiplesheets.xlsx");
    assert_fixture(&path);
    let sheet0 = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    let sheet1 = EasyExcel::read_dynamic_sync(&path)
        .sheet(1usize)
        .do_read_sync()
        .unwrap();
    assert!(!sheet0.is_empty());
    assert!(!sheet1.is_empty());
}

/// Ensure DynamicRow still works as no-model temp read path.
#[test]
fn temp_no_model_dynamic_row() {
    let path = fixture("demo/demo.csv");
    assert_fixture(&path);
    let rows: Vec<DynamicRow> = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
    assert!(!rows.is_empty());
}
