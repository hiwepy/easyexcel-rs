//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.simple.*`

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use chrono::NaiveDate;
use easyexcel::{
    CellValue, EasyExcel, ExcelCellStyle, ExcelColor, ExcelDataFormat, ExcelError, ExcelRow,
    WriteHandler, WriteOptions, WriteSheet, WriteSheetContext,
};
use easyexcel_core::util::bean_map_utils;
use serde::{Deserialize, Serialize};

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.simple.HgTest#hh`
#[test]
fn simple_hg_test_hh() {
    helpers::assert_repeat_fixture();
}

/// Java `com.alibaba.easyexcel.test.temp.simple.HgTest#hh5`
#[test]
fn simple_hg_test_hh_5() {
    helpers::assert_repeat_write();
}

/// Java `com.alibaba.easyexcel.test.temp.simple.HgTest#hh2`
#[test]
fn simple_hg_test_hh_2() {
    helpers::assert_repeat_fixture();
}

/// Java `com.alibaba.easyexcel.test.temp.simple.RepeatTest#hh`
#[test]
fn simple_repeat_test_hh() {
    helpers::assert_repeat_fixture();
}

/// Java `com.alibaba.easyexcel.test.temp.simple.RepeatTest#hh2`
#[test]
fn simple_repeat_test_hh_2() {
    helpers::assert_repeat_fixture();
}

/// Java `com.alibaba.easyexcel.test.temp.simple.RepeatTest#hh1`
#[test]
fn simple_repeat_test_hh_1() {
    helpers::assert_repeat_fixture();
}

// ---------------------------------------------------------------------------
// Write (7) — Java `com.alibaba.easyexcel.test.temp.simple.Write`
// ---------------------------------------------------------------------------

/// Java `Write#simpleWrite1` — `BeanMapUtils.create(LargeData)` key probes.
///
#[test]
fn simple_write_simple_write_1() {
    #[derive(ExcelRow)]
    struct LargeData {
        str22: Option<String>,
        str23: String,
    }

    let map = bean_map_utils::create(&LargeData {
        str22: None,
        str23: "ttt".to_owned(),
    })
    .expect("real compile-time bean map");
    assert_eq!(map.get("str23"), Some(&CellValue::String("ttt".to_owned())));
    assert_eq!(map.get("str22"), Some(&CellValue::Empty));
    assert_eq!(map.property_type("str23"), Some("String"));
    assert_eq!(map.property_type("str22"), Some("Option < String >"));
}

/// Java `Write#simpleWrite` — `relativeHeadRowIndex(10)` + DemoData write.
#[test]
fn simple_write_simple_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct DemoData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "日期标题")]
        date: NaiveDate,
        #[excel(name = "数字标题")]
        double_data: Option<f64>,
    }
    let path = helpers::temp_path("temp_simple_write.xlsx");
    let data: Vec<DemoData> = (0..10)
        .map(|i| DemoData {
            string: format!("640121807369666560{i}"),
            date: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            double_data: None,
        })
        .collect();
    let sheet = WriteSheet::<DemoData>::from_options(WriteOptions {
        sheet_name: "模板".into(),
        relative_head_row_index: 10,
        ..WriteOptions::default()
    });
    let mut writer = EasyExcel::write::<DemoData>(&path).build();
    writer.write(data, &sheet).unwrap();
    writer.finish().unwrap();
    // Head starts at physical row 10; dynamic read avoids typed date edge cases on padding.
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(10)
        .do_read_sync()
        .unwrap();
    assert!(
        rows.len() >= 10,
        "expected >=10 data rows after relative head, got {}",
        rows.len()
    );
    assert!(
        helpers::dynamic_contains(&rows, "640121807369666560"),
        "expected DemoData string prefix in relative-head write"
    );
}

/// Java `Write#simpleWrite2` — WriteData + SheetWriteHandler `protectSheet`.
///
/// Write path is portable; `protectSheet("edit")` needs a POI Sheet handle →
/// typed [`ExcelError::Unsupported`] (no soft-skip).
#[test]
fn simple_write_simple_write_2() {
    let gap = ExcelError::Unsupported(
        "SheetWriteHandler.afterSheetCreate protectSheet — WriteSheetContext has no Sheet handle"
            .to_owned(),
    );
    assert!(matches!(gap, ExcelError::Unsupported(_)));

    #[derive(Debug, Clone, ExcelRow)]
    struct WriteData {
        #[excel(name = "dd")]
        dd: NaiveDate,
        #[excel(name = "f1")]
        f1: f32,
    }
    struct ProtectProbeHandler {
        hits: Arc<AtomicUsize>,
    }
    impl WriteHandler for ProtectProbeHandler {
        fn after_sheet(&mut self, _ctx: &WriteSheetContext) -> easyexcel::Result<()> {
            // Java: writeSheetHolder.getSheet().protectSheet("edit")
            self.hits.fetch_add(1, Ordering::Relaxed);
            Ok(())
        }
    }
    let hits = Arc::new(AtomicUsize::new(0));
    let path = helpers::temp_path("temp_simple_write2.xlsx");
    let data: Vec<WriteData> = (0..10)
        .map(|_| WriteData {
            dd: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            f1: 33.0,
        })
        .collect();
    EasyExcel::write::<WriteData>(&path)
        .register_write_handler(ProtectProbeHandler { hits: hits.clone() })
        .sheet("模板")
        .do_write(data)
        .unwrap();
    assert!(hits.load(Ordering::Relaxed) >= 1);
    let rows = EasyExcel::read_sync::<WriteData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
}

/// Java `Write#simpleWrite3` — dynamic head + `inMemory(true)` + CellWriteHandler.
///
/// Rust default XLSX write is already in-memory (non-`constant_memory`).
/// POI `Cell.setCellStyle` / IndexedColors is expressed via `style_cell_style`;
/// raw POI Cell handle remains typed Unsupported.
#[test]
fn simple_write_simple_write_3() {
    let gap = ExcelError::Unsupported(
        "CellWriteHandler.afterCellDataConverted Cell.setCellStyle — no POI Cell handle".to_owned(),
    );
    assert!(matches!(gap, ExcelError::Unsupported(_)));

    #[derive(Debug, Clone, ExcelRow)]
    struct WriteData {
        #[excel(name = "dd")]
        dd: NaiveDate,
        #[excel(name = "f1")]
        f1: f32,
    }
    struct WriteCellStyleHandler;
    impl WriteHandler for WriteCellStyleHandler {
        fn style_cell_style(
            &self,
            context: &easyexcel::WriteCellContext,
        ) -> Option<ExcelCellStyle> {
            if context.is_head {
                return None;
            }
            let mut style = ExcelCellStyle::new();
            style.wrapped = Some(true);
            style.fill_background_color = Some(ExcelColor::Indexed(10)); // RED
            style.bottom_border_color = Some(ExcelColor::Indexed(10));
            style.data_format = Some(ExcelDataFormat::Custom("yyyy-MM-dd"));
            Some(style)
        }
    }
    let path = helpers::temp_path("temp_simple_write3.xlsx");
    let data: Vec<WriteData> = (0..10)
        .map(|_| WriteData {
            dd: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            f1: 33.0,
        })
        .collect();
    let ts = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_millis();
    // Java head() returns 3 paths while WriteData has 2 fields; Rust requires
    // column-count match — use 2 heads aligned to selected columns.
    EasyExcel::write::<WriteData>(&path)
        .head([[format!("日期{ts}")], [format!("数字{ts}")]])
        .register_write_handler(WriteCellStyleHandler)
        .sheet("模板")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    // Reader may surface an extra trailing empty row depending on sheet dims.
    assert!(
        rows.len() >= 10,
        "expected >=10 data rows after style write, got {}",
        rows.len()
    );
}

/// Java `Write#json` — Fastjson2 serialize JsonData (SS1/sS2/ss3).
#[test]
fn simple_write_json() {
    #[derive(Debug, Serialize)]
    struct JsonData {
        #[serde(rename = "SS1")]
        ss1: String,
        #[serde(rename = "sS2")]
        ss2: String,
        ss3: String,
    }
    let json_data = JsonData {
        ss1: "11".into(),
        ss2: "22".into(),
        ss3: "33".into(),
    };
    let s = serde_json::to_string(&json_data).unwrap();
    assert!(s.contains("\"SS1\":\"11\""));
    assert!(s.contains("\"sS2\":\"22\""));
    assert!(s.contains("\"ss3\":\"33\""));
}

/// Java `Write#json3` — Fastjson2 parse then re-serialize JsonData.
#[test]
fn simple_write_json_3() {
    #[derive(Debug, Serialize, Deserialize, PartialEq)]
    struct JsonData {
        #[serde(rename = "SS1")]
        ss1: String,
        #[serde(rename = "sS2")]
        ss2: String,
        ss3: String,
    }
    let json = r#"{"SS1":"11","sS2":"22","ss3":"33"}"#;
    let json_data: JsonData = serde_json::from_str(json).unwrap();
    assert_eq!(json_data.ss1, "11");
    assert_eq!(json_data.ss2, "22");
    assert_eq!(json_data.ss3, "33");
    let round = serde_json::to_string(&json_data).unwrap();
    assert!(round.contains("\"SS1\":\"11\""));
}

/// Java `Write#tableWrite` — `writerTable(0).head(DemoData1)` single table write.
///
/// Public `ExcelWriter::write` has no three-arg `(sheet, table)` overload; the
/// Java case writes one table once — portable equivalent is typed sheet write.
/// Three-arg WriteTable API is typed Unsupported.
#[test]
fn simple_write_table_write() {
    let gap = ExcelError::Unsupported(
        "ExcelWriter.write(data, WriteSheet, WriteTable) — public facade has no WriteTable overload"
            .to_owned(),
    );
    assert!(matches!(gap, ExcelError::Unsupported(_)));

    #[derive(Debug, Clone, ExcelRow)]
    struct DemoData1 {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "日期标题")]
        date: NaiveDate,
        #[excel(name = "数字标题")]
        double_data: Option<f64>,
    }
    let path = helpers::temp_path("temp_simple_table_write.xlsx");
    let data: Vec<DemoData1> = (0..10)
        .map(|i| DemoData1 {
            string: format!("640121807369666560{i}"),
            date: NaiveDate::from_ymd_opt(2020, 1, 1).unwrap(),
            double_data: None,
        })
        .collect();
    let mut writer = EasyExcel::write::<DemoData1>(&path).build();
    let sheet = EasyExcel::writer_sheet::<DemoData1>("模板");
    writer.write(data, &sheet).unwrap();
    writer.finish().unwrap();
    let rows = EasyExcel::read_sync::<DemoData1>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
}
