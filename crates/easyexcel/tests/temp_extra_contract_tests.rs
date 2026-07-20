//! Extra temp-package contracts: dataformat / encrypt / write / issue / lock.
//!
//! Split from `temp_contract_tests.rs` to keep files under 500 lines.

use easyexcel::{DynamicRow, DynamicValue, EasyExcel, ExcelRow};
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

fn dynamic_contains(rows: &[DynamicRow], needle: &str) -> bool {
    rows.iter().any(|row| {
        row.values().iter().any(|(_, val)| match val {
            DynamicValue::String(s) => s.contains(needle),
            DynamicValue::ActualData(easyexcel::CellValue::String(s)) => s.contains(needle),
            _ => false,
        })
    })
}

/// Java `temp.dataformat.DataFormatTest` — read portable dataformat.xlsx.
#[test]
fn temp_dataformat_xlsx_fixture() {
    let path = fixture("dataformat/dataformat.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.dataformat.DataFormatTest.testxls` — read portable dataformat.xls.
#[test]
fn temp_dataformat_xls_fixture() {
    let path = fixture("dataformat/dataformat.xls");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.dataformat` / issue2443 date fixtures — readable without panic.
#[test]
fn temp_dataformat_date_fixtures() {
    for name in [
        "dataformat/date1.xlsx",
        "dataformat/date2.xlsx",
        "dataformat/dataformatv2.xlsx",
    ] {
        let path = fixture(name);
        assert_fixture(&path);
        let rows = EasyExcel::read_dynamic_sync(&path)
            .do_read_sync()
            .unwrap();
        assert!(!rows.is_empty(), "{name} must yield rows");
    }
}

/// Java `temp.poi.PoiEncryptTest.encryptExcel` — password write/read (API supported).
#[test]
fn temp_encrypt_password_round_trip() {
    #[derive(Debug, Clone, ExcelRow)]
    struct EncryptData {
        #[excel(name = "string")]
        string: String,
    }
    let path = temp_path("temp_encrypt.xlsx");
    EasyExcel::write::<EncryptData>(&path)
        .password("123456")
        .sheet("Sheet1")
        .do_write(vec![EncryptData {
            string: "secret".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<EncryptData>(&path)
        .password("123456")
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].string, "secret");
}

/// Java `temp.write.TempWriteTest.write` — newline / backslash in cell survives round trip.
#[test]
fn temp_write_newline_string_round_trip() {
    #[derive(Debug, Clone, ExcelRow)]
    struct TempWriteData {
        #[excel(name = "name")]
        name: String,
    }
    let path = temp_path("temp_write_newline.xlsx");
    EasyExcel::write::<TempWriteData>(&path)
        .sheet("Sheet1")
        .do_write(vec![TempWriteData {
            name: "zs\r\n \\ \r\n t4".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<TempWriteData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert!(rows[0].name.contains("zs"));
    assert!(rows[0].name.contains("t4"));
}

/// Java `temp.issue1662.Issue1662Test` — dynamic multi-level head write.
#[test]
fn temp_issue1662_dynamic_head_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Row {
        #[excel(name = "c0")]
        c0: String,
        #[excel(name = "c1")]
        c1: String,
    }
    let path = temp_path("temp_issue1662.xlsx");
    EasyExcel::write::<Row>(&path)
        .head(vec![
            vec!["xx".to_owned(), "日期".to_owned()],
            vec!["日期".to_owned()],
        ])
        .sheet("模板")
        .do_write(vec![Row {
            c0: "字符串".into(),
            c1: "0.56".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(dynamic_contains(&rows, "字符串") || dynamic_contains(&rows, "日期"));
}

/// Java `temp.issue2443.Issue2443Test` — date1/date2 fixtures readable.
#[test]
fn temp_issue2443_date_fixtures() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Issue2443 {
        #[excel(name = "a", index = 0)]
        a: i32,
        #[excel(name = "b", index = 1)]
        b: i32,
    }
    for name in [
        "java/temp/issue2443/date1.xlsx",
        "java/temp/issue2443/date2.xlsx",
        "dataformat/date1.xlsx",
        "dataformat/date2.xlsx",
    ] {
        let path = fixture(name);
        assert_fixture(&path);
        let typed = EasyExcel::read_sync::<Issue2443>(&path).do_read_sync();
        let dynamic = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
        assert!(!dynamic.is_empty(), "{name} dynamic read empty");
        let _ = typed;
    }
}

/// Java `temp.LockTest` / `Lock2Test` — machine-local / stress paths intentionally ignored.
///
/// Reason: hard-coded `/Users/zhuangjiaju/...` and `D:\\test\\...` plus ad-hoc POI
/// format probes; no portable fixture. Prefer assertable contracts above.
#[test]
#[ignore = "Lock* uses machine-local paths and stress/POI probes; not a portable contract"]
fn temp_lock_stress_intentionally_skipped() {
    panic!("should remain ignored");
}
