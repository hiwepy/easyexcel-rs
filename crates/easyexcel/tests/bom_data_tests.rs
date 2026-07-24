//! BomDataTest parity — mirrors Java
//! `com.alibaba.easyexcel.test.core.bom.BomDataTest`.
//!
//! Java methods:
//! - `t01ReadCsv` — read `bom/no_bom.csv` and `bom/office_bom.csv`
//! - `t02ReadAndWriteCsv` — write/read with charset + withBom variants
//!
//! Assertions match Java: head `"姓名"`, 10 rows, first name `"姓名0"`, age `20`.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use easyexcel::{
    AnalysisContext, CsvCharset, EasyExcel, ExcelRow, PageReadListener, ReadListener, Result,
};

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn temp_path(name: &str) -> std::path::PathBuf {
    let dir = tempfile::tempdir().unwrap();
    dir.keep().join(name)
}

/// Java: `com.alibaba.easyexcel.test.core.bom.BomData`
#[derive(Debug, Clone, ExcelRow)]
struct BomData {
    #[excel(name = "姓名")]
    name: String,
    #[excel(name = "年纪")]
    age: i64,
}

/// Java `BomDataTest.data()` — 10 rows, name=`姓名{i}`, age=20.
fn bom_data() -> Vec<BomData> {
    (0..10)
        .map(|i| BomData {
            name: format!("姓名{i}"),
            age: 20,
        })
        .collect()
}

struct BomReadListener {
    heads: Vec<String>,
    rows: Vec<BomData>,
}

impl ReadListener<BomData> for BomReadListener {
    /// Java `invokeHead` — assert column 0 header is `"姓名"`.
    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        // Rust head map is name → index; Java used index → ReadCellData.
        let mut by_index: Vec<(usize, String)> =
            head.iter().map(|(k, v)| (*v, k.clone())).collect();
        by_index.sort_by_key(|(idx, _)| *idx);
        if let Some((_, name)) = by_index.first() {
            self.heads.push(name.clone());
        }
        Ok(())
    }

    fn invoke(&mut self, data: BomData, _context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        Ok(())
    }

    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
        assert_eq!(self.heads.first().map(String::as_str), Some("姓名"));
        assert_eq!(self.rows.len(), 10);
        assert_eq!(self.rows[0].name, "姓名0");
        assert_eq!(self.rows[0].age, 20);
        Ok(())
    }
}

/// Java `BomDataTest.readCsv`.
fn assert_read_csv(path: &std::path::Path) {
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
    let listener = BomReadListener {
        heads: Vec::new(),
        rows: Vec::new(),
    };
    EasyExcel::read::<BomData, _>(path, listener)
        .sheet(0usize)
        .do_read()
        .unwrap();
}

/// Java `BomDataTest.t01ReadCsv`.
#[test]
fn bom_t01_read_csv() {
    // Java: readCsv(no_bom) + readCsv(office_bom)
    assert_read_csv(&fixture("bom/no_bom.csv"));
    assert_read_csv(&fixture("bom/office_bom.csv"));
}

/// Java `BomDataTest.readAndWriteCsv` — charset + optional withBom.
fn assert_read_and_write_csv(
    path: &std::path::Path,
    charset: Option<&str>,
    with_bom: Option<bool>,
) {
    let mut writer = EasyExcel::write::<BomData>(path);
    if let Some(cs) = charset {
        writer = writer.charset(CsvCharset::new(cs));
    }
    if let Some(bom) = with_bom {
        writer = writer.with_bom(bom);
    }
    writer.sheet("Sheet1").do_write(bom_data()).unwrap();

    let mut reader = EasyExcel::read_sync::<BomData>(path);
    if let Some(cs) = charset {
        reader = reader.charset(CsvCharset::new(cs));
    }
    let rows = reader.do_read_sync().unwrap();
    assert_eq!(
        rows.len(),
        10,
        "Java asserts dataList.size()==10 for {}",
        path.display()
    );
    assert_eq!(rows[0].name, "姓名0");
    assert_eq!(rows[0].age, 20);
}

/// Java `BomDataTest.t02ReadAndWriteCsv`.
#[test]
fn bom_t02_read_and_write_csv() {
    assert_read_and_write_csv(&temp_path("bom_default.csv"), None, None);
    assert_read_and_write_csv(&temp_path("bom_utf_8.csv"), Some("UTF-8"), None);
    assert_read_and_write_csv(&temp_path("bom_utf_8_lower_case.csv"), Some("utf-8"), None);
    assert_read_and_write_csv(&temp_path("bom_gbk.csv"), Some("GBK"), None);
    assert_read_and_write_csv(&temp_path("bom_gbk_lower_case.csv"), Some("gbk"), None);
    assert_read_and_write_csv(&temp_path("bom_utf_16be.csv"), Some("UTF-16BE"), None);
    assert_read_and_write_csv(
        &temp_path("bom_utf_8_not_with_bom.csv"),
        Some("UTF-8"),
        Some(false),
    );
}

/// PageReadListener path (Java demos) should also see 10 rows for office_bom.
#[test]
fn bom_page_read_listener_office_bom() {
    let path = fixture("bom/office_bom.csv");
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
    let total = Arc::new(Mutex::new(0usize));
    let total_cb = Arc::clone(&total);
    let listener = PageReadListener::new(100, move |batch: Vec<BomData>, _ctx| {
        *total_cb.lock().unwrap() += batch.len();
        Ok(())
    });
    EasyExcel::read::<BomData, _>(&path, listener)
        .sheet(0usize)
        .do_read()
        .unwrap();
    assert_eq!(*total.lock().unwrap(), 10);
}
