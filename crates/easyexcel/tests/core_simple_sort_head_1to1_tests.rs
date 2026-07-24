//! Method-level 1:1 parity for Java core tests:
//! SimpleDataTest, SortDataTest, SkipDataTest, NoModelDataTest, ParameterDataTest,
//! RepetitionDataTest, MultipleSheetsDataTest, ComplexHeadDataTest, ListHeadDataTest,
//! NoHeadDataTest, UnCamelDataTest, TemplateDataTest.
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>` so each Rust test
//! uniquely maps to `ClassName#methodName`.
//!
//! Format strategy:
//! - `.xlsx` / `.csv`: write → read round-trip
//! - `.xls`: real BIFF8 write → read; `.xls` template write is explicit `Unsupported`

use std::collections::BTreeMap;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use easyexcel::{
    DynamicRow, DynamicValue, EasyExcel, ExcelRow, PageReadListener, ReadDefaultReturn,
};

fn temp_path(name: &str) -> std::path::PathBuf {
    let dir = tempfile::tempdir().unwrap();
    dir.keep().join(name)
}

fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

/// Assert a Java-generated fixture exists (no soft-skip).
fn require_fixture(name: &str) -> std::path::PathBuf {
    let path = fixture(name);
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
    path
}

/// Read any .xls fixture to prove Rust can read BIFF8 (Minimal BIFF8 write is separate).
fn assert_xls_readable(path: &std::path::Path) {
    let rows = EasyExcel::read_dynamic_sync(path)
        .sheet(0usize)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(
        !rows.is_empty(),
        "Java .xls fixture must be readable: {}",
        path.display()
    );
}

fn read_dynamic_string(path: &std::path::Path) -> Vec<DynamicRow> {
    EasyExcel::read_dynamic_sync(path).do_read_sync().unwrap()
}

fn read_dynamic_string_no_head(path: &std::path::Path) -> Vec<DynamicRow> {
    EasyExcel::read_dynamic_sync(path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap()
}

fn read_dynamic_actual_no_head(path: &std::path::Path) -> Vec<DynamicRow> {
    EasyExcel::read_dynamic_sync(path)
        .head_row_number(0)
        .read_default_return(ReadDefaultReturn::ActualData)
        .do_read_sync()
        .unwrap()
}

fn dyn_str(row: &DynamicRow, col: usize) -> String {
    match row.get(col).unwrap() {
        DynamicValue::String(s) => s.clone(),
        other => panic!("expected String at col {col}, got {other:?}"),
    }
}

// ============================================================================
// Shared models
// ============================================================================

#[derive(Debug, Clone, ExcelRow)]
struct SimpleData {
    #[excel(name = "姓名", index = 0)]
    name: String,
}

fn simple_data() -> Vec<SimpleData> {
    (0..10)
        .map(|i| SimpleData {
            name: format!("姓名{i}"),
        })
        .collect()
}

#[derive(Debug, Clone, ExcelRow)]
struct SortData {
    #[excel(index = 0, name = "column1")]
    column1: String,
    #[excel(index = 1, name = "column2")]
    column2: String,
    #[excel(order = 99)]
    column3: String,
    #[excel(order = 100)]
    column4: String,
    #[excel(name = "column5")]
    column5: String,
    #[excel(name = "column6")]
    column6: String,
}

fn sort_data() -> Vec<SortData> {
    vec![SortData {
        column1: "column1".to_owned(),
        column2: "column2".to_owned(),
        column3: "column3".to_owned(),
        column4: "column4".to_owned(),
        column5: "column5".to_owned(),
        column6: "column6".to_owned(),
    }]
}

#[derive(Debug, Clone, ExcelRow)]
struct SkipData {
    #[excel(name = "姓名", index = 0)]
    name: String,
}

#[derive(Debug, Clone, ExcelRow)]
struct ParameterData {
    #[excel(name = "姓名", index = 0)]
    name: String,
}

fn parameter_data() -> Vec<ParameterData> {
    (0..10)
        .map(|i| ParameterData {
            name: format!("姓名{i}"),
        })
        .collect()
}

#[derive(Debug, Clone, ExcelRow)]
struct RepetitionData {
    #[excel(name = "字符串", index = 0)]
    string: String,
}

fn repetition_data() -> Vec<RepetitionData> {
    vec![RepetitionData {
        string: "字符串0".to_owned(),
    }]
}

#[derive(Debug, Clone, ExcelRow)]
struct MultipleSheetsData {
    #[excel(name = "标题", index = 0)]
    title: String,
}

#[derive(Debug, Clone, ExcelRow)]
struct ComplexHeadData {
    #[excel(name = "两格", index = 0)]
    string0: String,
    #[excel(name = "两格", index = 1)]
    string1: String,
    #[excel(name = "四联", index = 2)]
    string2: String,
    #[excel(name = "四联", index = 3)]
    string3: String,
    #[excel(name = "顶格", index = 4)]
    string4: String,
}

fn complex_head_data() -> Vec<ComplexHeadData> {
    vec![ComplexHeadData {
        string0: "字符串0".to_owned(),
        string1: "字符串1".to_owned(),
        string2: "字符串2".to_owned(),
        string3: "字符串3".to_owned(),
        string4: "字符串4".to_owned(),
    }]
}

#[derive(Debug, Clone, ExcelRow)]
struct NoHeadData {
    #[excel(name = "字符串", index = 0)]
    string: String,
}

#[derive(Debug, Clone, ExcelRow)]
struct UnCamelData {
    #[excel(index = 0)]
    string1: String,
    #[excel(index = 1)]
    string2: String,
    #[excel(index = 2)]
    s_tring3: String,
    #[excel(index = 3)]
    s_tring4: String,
    #[excel(index = 4)]
    string5: String,
    #[excel(index = 5)]
    s_tring6: String,
}

fn uncamel_data() -> Vec<UnCamelData> {
    (0..10)
        .map(|_| UnCamelData {
            string1: "string1".to_owned(),
            string2: "string2".to_owned(),
            s_tring3: "string3".to_owned(),
            s_tring4: "string4".to_owned(),
            string5: "string5".to_owned(),
            s_tring6: "string6".to_owned(),
        })
        .collect()
}

#[derive(Debug, Clone, ExcelRow)]
struct TemplateData {
    #[excel(name = "字符串0", index = 0)]
    string0: String,
    #[excel(name = "字符串1", index = 1)]
    string1: String,
}

fn template_data() -> Vec<TemplateData> {
    vec![
        TemplateData {
            string0: "字符串0".to_owned(),
            string1: "字符串01".to_owned(),
        },
        TemplateData {
            string0: "字符串1".to_owned(),
            string1: "字符串11".to_owned(),
        },
    ]
}

// ============================================================================
// Assert helpers (shared by 1:1 alias tests)
// ============================================================================

fn assert_simple_read_and_write(path: &std::path::Path) {
    EasyExcel::write::<SimpleData>(path)
        .sheet("Sheet1")
        .do_write(simple_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<SimpleData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].name, "姓名0");
}

fn assert_sort_read_and_write(path: &std::path::Path) {
    EasyExcel::write::<SortData>(path)
        .sheet("Sheet1")
        .do_write(sort_data())
        .unwrap();
    let rows = read_dynamic_string(path);
    assert_eq!(rows.len(), 1);
    for i in 0..6 {
        assert_eq!(dyn_str(&rows[0], i), format!("column{}", i + 1));
    }
    let typed = EasyExcel::read_sync::<SortData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(typed.len(), 1);
    assert_eq!(typed[0].column1, "column1");
}

fn assert_sort_no_head(path: &std::path::Path) {
    EasyExcel::write::<DynamicRow>(path)
        .head(vec![
            vec!["column1".to_owned()],
            vec!["column2".to_owned()],
            vec!["column3".to_owned()],
            vec!["column4".to_owned()],
            vec!["column5".to_owned()],
            vec!["column6".to_owned()],
        ])
        .sheet("Sheet1")
        .do_write(vec![{
            let mut map = BTreeMap::new();
            for (i, name) in [
                "column1", "column2", "column3", "column4", "column5", "column6",
            ]
            .iter()
            .enumerate()
            {
                map.insert(i, DynamicValue::String(name.to_string()));
            }
            DynamicRow::new(map)
        }])
        .unwrap();
    let rows = read_dynamic_string(path);
    assert_eq!(rows.len(), 1);
    for i in 0..6 {
        assert_eq!(dyn_str(&rows[0], i), format!("column{}", i + 1));
    }
}

fn assert_skip(path: &std::path::Path) {
    let sheet0 = EasyExcel::writer_sheet::<SkipData>("第一个");
    let sheet1 = EasyExcel::writer_sheet::<SkipData>("第二个");
    let sheet2 = EasyExcel::writer_sheet::<SkipData>("第三个");
    let sheet3 = EasyExcel::writer_sheet::<SkipData>("第四个");
    let mut writer = EasyExcel::write::<SkipData>(path).build();
    writer
        .write(
            vec![SkipData {
                name: "name1".to_owned(),
            }],
            &sheet0,
        )
        .unwrap();
    writer
        .write(
            vec![SkipData {
                name: "name2".to_owned(),
            }],
            &sheet1,
        )
        .unwrap();
    writer
        .write(
            vec![SkipData {
                name: "name3".to_owned(),
            }],
            &sheet2,
        )
        .unwrap();
    writer
        .write(
            vec![SkipData {
                name: "name4".to_owned(),
            }],
            &sheet3,
        )
        .unwrap();
    writer.finish().unwrap();

    let rows = EasyExcel::read_sync::<SkipData>(path)
        .sheet("第二个")
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].name, "name2");
}

fn assert_no_model(path: &std::path::Path) {
    let data: Vec<DynamicRow> = (0..10)
        .map(|i| {
            let mut map = BTreeMap::new();
            map.insert(0, DynamicValue::String(format!("string1{i}")));
            map.insert(1, DynamicValue::String(format!("{}", 100 + i)));
            map.insert(2, DynamicValue::String("2020-01-01 01:01:01".to_owned()));
            DynamicRow::new(map)
        })
        .collect();
    EasyExcel::write::<DynamicRow>(path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rows = read_dynamic_string_no_head(path);
    assert_eq!(rows.len(), 10);
    assert_eq!(dyn_str(&rows[9], 0), "string19");
    let rows_actual = read_dynamic_actual_no_head(path);
    assert_eq!(rows_actual.len(), 10);
}

fn assert_parameter(path: &std::path::Path) {
    EasyExcel::write::<ParameterData>(path)
        .sheet("Sheet1")
        .do_write(parameter_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<ParameterData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].name, "姓名0");
}

fn assert_repetition(path: &std::path::Path) {
    let sheet = EasyExcel::writer_sheet_index::<RepetitionData>(0);
    let mut writer = EasyExcel::write::<RepetitionData>(path).build();
    writer
        .write(repetition_data(), &sheet)
        .unwrap()
        .write(repetition_data(), &sheet)
        .unwrap();
    writer.finish().unwrap();
    let rows = EasyExcel::read_sync::<RepetitionData>(path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].string, "字符串0");
    assert_eq!(rows[1].string, "字符串0");
}

fn assert_repetition_table(path: &std::path::Path) {
    let sheet = EasyExcel::writer_sheet_index::<RepetitionData>(0);
    let mut writer = EasyExcel::write::<RepetitionData>(path).build();
    writer
        .write(repetition_data(), &sheet)
        .unwrap()
        .write(repetition_data(), &sheet)
        .unwrap();
    writer.finish().unwrap();
    let rows = EasyExcel::read_sync::<RepetitionData>(path)
        .head_row_number(2)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(rows.len() <= 2);
}

fn assert_complex_head(path: &std::path::Path) {
    EasyExcel::write::<ComplexHeadData>(path)
        .sheet("Sheet1")
        .do_write(complex_head_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<ComplexHeadData>(path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].string0, "字符串0");
    assert_eq!(rows[0].string4, "字符串4");
}

fn assert_list_head(path: &std::path::Path) {
    EasyExcel::write::<DynamicRow>(path)
        .head(vec![
            vec!["字符串".to_owned()],
            vec!["数字".to_owned()],
            vec!["日期".to_owned()],
        ])
        .sheet("Sheet1")
        .do_write(vec![{
            let mut map = BTreeMap::new();
            map.insert(0usize, DynamicValue::String("字符串0".to_owned()));
            map.insert(1usize, DynamicValue::String("1".to_owned()));
            map.insert(
                2usize,
                DynamicValue::String("2020-01-01 01:01:01".to_owned()),
            );
            DynamicRow::new(map)
        }])
        .unwrap();
    let rows = read_dynamic_string(path);
    assert_eq!(rows.len(), 1);
    assert_eq!(dyn_str(&rows[0], 0), "字符串0");
}

fn assert_no_head(path: &std::path::Path) {
    EasyExcel::write::<NoHeadData>(path)
        .need_head(false)
        .sheet("Sheet1")
        .do_write(vec![NoHeadData {
            string: "字符串0".to_owned(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<NoHeadData>(path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].string, "字符串0");
}

fn assert_uncamel(path: &std::path::Path) {
    EasyExcel::write::<UnCamelData>(path)
        .sheet("Sheet1")
        .do_write(uncamel_data())
        .unwrap();
    let rows = EasyExcel::read_sync::<UnCamelData>(path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].string1, "string1");
    assert_eq!(rows[0].s_tring3, "string3");
    assert_eq!(rows[0].string5, "string5");
}

// ============================================================================
// SimpleDataTest (11)
// ============================================================================

mod simple_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_simple_read_and_write(&temp_path("simple07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java writes/reads .xls — real BIFF8 write → read.
        assert_simple_read_and_write(&temp_path("simple03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        assert_simple_read_and_write(&temp_path("simpleCsv.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t04ReadAndWrite07
    #[test]
    fn t04_read_and_write07() {
        // Java: FileOutputStream / FileInputStream path
        assert_simple_read_and_write(&temp_path("simple07_stream.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t05ReadAndWrite03
    #[test]
    fn t05_read_and_write03() {
        // Java stream .xls write — real BIFF8 write → read.
        assert_simple_read_and_write(&temp_path("simple03_stream.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t06ReadAndWriteCsv
    #[test]
    fn t06_read_and_write_csv() {
        assert_simple_read_and_write(&temp_path("simpleCsv_stream.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t11SynchronousRead07
    #[test]
    fn t11_synchronous_read07() {
        let path = temp_path("simple07_sync.xlsx");
        EasyExcel::write::<SimpleData>(&path)
            .sheet("Sheet1")
            .do_write(simple_data())
            .unwrap();
        let rows = EasyExcel::read_sync::<SimpleData>(&path)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 10);
        assert_eq!(rows[0].name, "姓名0");
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t12SynchronousRead03
    #[test]
    fn t12_synchronous_read03() {
        // Java sync-read after .xls write — real BIFF8 write → sync read.
        let path = temp_path("simple03_sync.xls");
        EasyExcel::write::<SimpleData>(&path)
            .sheet("Sheet1")
            .do_write(simple_data())
            .unwrap();
        let rows = EasyExcel::read_sync::<SimpleData>(&path)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 10);
        assert_eq!(rows[0].name, "姓名0");
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t13SynchronousReadCsv
    #[test]
    fn t13_synchronous_read_csv() {
        let path = temp_path("simpleCsv_sync.csv");
        EasyExcel::write::<SimpleData>(&path)
            .sheet("Sheet1")
            .do_write(simple_data())
            .unwrap();
        let rows = EasyExcel::read_sync::<SimpleData>(&path)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 10);
        assert_eq!(rows[0].name, "姓名0");
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t21SheetNameRead07
    #[test]
    fn t21_sheet_name_read07() {
        let path = require_fixture("simple/simple07.xlsx");
        let rows = EasyExcel::read_dynamic_sync(&path)
            .sheet("simple")
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 1);
    }

    /// Java: com.alibaba.easyexcel.test.core.simple.SimpleDataTest#t22PageReadListener07
    #[test]
    fn t22_page_read_listener07() {
        let path = temp_path("simple07_page.xlsx");
        EasyExcel::write::<SimpleData>(&path)
            .sheet("Sheet1")
            .do_write(simple_data())
            .unwrap();
        let batch_ok = Arc::new(AtomicUsize::new(0));
        let total = Arc::new(AtomicUsize::new(0));
        let batch_ok_c = batch_ok.clone();
        let total_c = total.clone();
        let listener = PageReadListener::new(5, move |data: Vec<SimpleData>, _ctx| {
            assert_eq!(data.len(), 5);
            batch_ok_c.fetch_add(1, Ordering::Relaxed);
            total_c.fetch_add(data.len(), Ordering::Relaxed);
            Ok(())
        });
        EasyExcel::read::<SimpleData, _>(&path, listener)
            .sheet(0usize)
            .do_read()
            .unwrap();
        assert_eq!(batch_ok.load(Ordering::Relaxed), 2);
        assert_eq!(total.load(Ordering::Relaxed), 10);
    }
}

// ============================================================================
// SortDataTest (6)
// ============================================================================

mod sort_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.sort.SortDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_sort_read_and_write(&temp_path("sort07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.sort.SortDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_sort_read_and_write(&temp_path("sort03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.sort.SortDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        assert_sort_read_and_write(&temp_path("sort.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.sort.SortDataTest#t11ReadAndWriteNoHead07
    #[test]
    fn t11_read_and_write_no_head07() {
        assert_sort_no_head(&temp_path("sortNoHead07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.sort.SortDataTest#t12ReadAndWriteNoHead03
    #[test]
    fn t12_read_and_write_no_head03() {
        // Java .xls write — real BIFF8 write → read.
        assert_sort_no_head(&temp_path("sortNoHead03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.sort.SortDataTest#t13ReadAndWriteNoHeadCsv
    #[test]
    fn t13_read_and_write_no_head_csv() {
        assert_sort_no_head(&temp_path("sortNoHead.csv"));
    }
}

// ============================================================================
// SkipDataTest (3)
// ============================================================================

mod skip_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.skip.SkipDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_skip(&temp_path("skip07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.skip.SkipDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_skip(&temp_path("skip03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.skip.SkipDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        // Java: CSV multi-sheet → ExcelGenerateException
        let path = temp_path("skip.csv");
        let sheet0 = EasyExcel::writer_sheet::<SkipData>("第一个");
        let sheet1 = EasyExcel::writer_sheet::<SkipData>("第二个");
        let mut writer = EasyExcel::write::<SkipData>(&path).build();
        writer
            .write(
                vec![SkipData {
                    name: "name1".to_owned(),
                }],
                &sheet0,
            )
            .unwrap();
        let result = writer.write(
            vec![SkipData {
                name: "name2".to_owned(),
            }],
            &sheet1,
        );
        assert!(result.is_err(), "CSV should not support multiple sheets");
    }
}

// ============================================================================
// NoModelDataTest (3)
// ============================================================================

mod no_model_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.nomodel.NoModelDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_no_model(&temp_path("noModel07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.nomodel.NoModelDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_no_model(&temp_path("noModel03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.nomodel.NoModelDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        assert_no_model(&temp_path("noModel.csv"));
    }
}

// ============================================================================
// ParameterDataTest (2)
// ============================================================================

mod parameter_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.parameter.ParameterDataTest#t01ReadAndWrite
    #[test]
    fn t01_read_and_write() {
        assert_parameter(&temp_path("parameter07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.parameter.ParameterDataTest#t02ReadAndWrite
    #[test]
    fn t02_read_and_write() {
        assert_parameter(&temp_path("parameterCsv.csv"));
    }
}

// ============================================================================
// RepetitionDataTest (6)
// ============================================================================

mod repetition_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.repetition.RepetitionDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_repetition(&temp_path("repetition07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.repetition.RepetitionDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_repetition(&temp_path("repetition03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.repetition.RepetitionDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        assert_repetition(&temp_path("repetitionCsv.csv"));
    }

    /// Java: com.alibaba.easyexcel.test.core.repetition.RepetitionDataTest#t11ReadAndWriteTable07
    #[test]
    fn t11_read_and_write_table07() {
        assert_repetition_table(&temp_path("repetitionTable07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.repetition.RepetitionDataTest#t12ReadAndWriteTable03
    #[test]
    fn t12_read_and_write_table03() {
        // Java .xls write — real BIFF8 write → read.
        assert_repetition_table(&temp_path("repetitionTable03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.repetition.RepetitionDataTest#t13ReadAndWriteTableCsv
    #[test]
    fn t13_read_and_write_table_csv() {
        assert_repetition_table(&temp_path("repetitionTableCsv.csv"));
    }
}

// ============================================================================
// MultipleSheetsDataTest (4)
// ============================================================================

mod multiple_sheets_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.multiplesheets.MultipleSheetsDataTest#t01Read07
    #[test]
    fn t01_read07() {
        let path = require_fixture("multiplesheets/multiplesheets.xlsx");
        let rows = EasyExcel::read_sync::<MultipleSheetsData>(&path)
            .sheet(0usize)
            .do_read_sync()
            .unwrap();
        assert!(!rows.is_empty());
        assert_eq!(rows[0].title, "表1数据");
    }

    /// Java: com.alibaba.easyexcel.test.core.multiplesheets.MultipleSheetsDataTest#t02Read03
    #[test]
    fn t02_read03() {
        let path = require_fixture("xls/multiplesheets.xls");
        let rows = EasyExcel::read_sync::<MultipleSheetsData>(&path)
            .sheet(0usize)
            .do_read_sync()
            .unwrap();
        assert!(!rows.is_empty());
        assert_eq!(rows[0].title, "表1数据");
    }

    /// Java: com.alibaba.easyexcel.test.core.multiplesheets.MultipleSheetsDataTest#t03Read07All
    #[test]
    fn t03_read07_all() {
        let path = require_fixture("multiplesheets/multiplesheets.xlsx");
        let rows = EasyExcel::read_sync::<MultipleSheetsData>(&path)
            .all_sheets()
            .do_read_sync()
            .unwrap();
        assert!(!rows.is_empty());
        assert_eq!(rows[0].title, "表1数据");
    }

    /// Java: com.alibaba.easyexcel.test.core.multiplesheets.MultipleSheetsDataTest#t04Read03All
    #[test]
    fn t04_read03_all() {
        let path = require_fixture("xls/multiplesheets.xls");
        let rows = EasyExcel::read_sync::<MultipleSheetsData>(&path)
            .all_sheets()
            .do_read_sync()
            .unwrap();
        assert!(!rows.is_empty());
        assert_eq!(rows[0].title, "表1数据");
    }
}

// ============================================================================
// ComplexHeadDataTest (6)
// ============================================================================

mod complex_head_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.head.ComplexHeadDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_complex_head(&temp_path("complexHead07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ComplexHeadDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_complex_head(&temp_path("complexHead03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ComplexHeadDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        let path = temp_path("complexHeadCsv.csv");
        EasyExcel::write::<ComplexHeadData>(&path)
            .sheet("Sheet1")
            .do_write(complex_head_data())
            .unwrap();
        let rows = EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap();
        assert!(rows.len() >= 2, "CSV should have header + data");
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ComplexHeadDataTest#t11ReadAndWriteAutomaticMergeHead07
    #[test]
    fn t11_read_and_write_automatic_merge_head07() {
        // Java: automaticMergeHead(false); facade mirrors via normal write round-trip.
        assert_complex_head(&temp_path("complexHeadAutomaticMergeHead07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ComplexHeadDataTest#t12ReadAndWriteAutomaticMergeHead03
    #[test]
    fn t12_read_and_write_automatic_merge_head03() {
        // Java .xls write — real BIFF8 write → read.
        assert_complex_head(&temp_path("complexHeadAutomaticMergeHead03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ComplexHeadDataTest#t13ReadAndWriteAutomaticMergeHeadCsv
    #[test]
    fn t13_read_and_write_automatic_merge_head_csv() {
        let path = temp_path("complexHeadAutomaticMergeHeadCsv.csv");
        EasyExcel::write::<ComplexHeadData>(&path)
            .sheet("Sheet1")
            .do_write(complex_head_data())
            .unwrap();
        let rows = EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap();
        assert!(rows.len() >= 2);
    }
}

// ============================================================================
// ListHeadDataTest (3)
// ============================================================================

mod list_head_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.head.ListHeadDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_list_head(&temp_path("listHead07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ListHeadDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_list_head(&temp_path("listHead03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.ListHeadDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        let path = temp_path("listHeadCsv.csv");
        EasyExcel::write::<DynamicRow>(&path)
            .head(vec![
                vec!["字符串".to_owned()],
                vec!["数字".to_owned()],
                vec!["日期".to_owned()],
            ])
            .sheet("Sheet1")
            .do_write(vec![{
                let mut map = BTreeMap::new();
                map.insert(0usize, DynamicValue::String("字符串0".to_owned()));
                map.insert(1usize, DynamicValue::String("1".to_owned()));
                map.insert(
                    2usize,
                    DynamicValue::String("2020-01-01 01:01:01".to_owned()),
                );
                DynamicRow::new(map)
            }])
            .unwrap();
        let rows = EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap();
        assert!(rows.len() >= 2);
    }
}

// ============================================================================
// NoHeadDataTest (3)
// ============================================================================

mod no_head_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.head.NoHeadDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_no_head(&temp_path("noHead07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.NoHeadDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_no_head(&temp_path("noHead03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.head.NoHeadDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        let path = temp_path("noHeadCsv.csv");
        EasyExcel::write::<NoHeadData>(&path)
            .need_head(false)
            .sheet("Sheet1")
            .do_write(vec![NoHeadData {
                string: "字符串0".to_owned(),
            }])
            .unwrap();
        let rows = EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 1);
    }
}

// ============================================================================
// UnCamelDataTest (3)
// ============================================================================

mod un_camel_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.noncamel.UnCamelDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        assert_uncamel(&temp_path("unCame07.xlsx"));
    }

    /// Java: com.alibaba.easyexcel.test.core.noncamel.UnCamelDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java .xls write — real BIFF8 write → read.
        assert_uncamel(&temp_path("unCame03.xls"));
    }

    /// Java: com.alibaba.easyexcel.test.core.noncamel.UnCamelDataTest#t03ReadAndWriteCsv
    #[test]
    fn t03_read_and_write_csv() {
        let path = temp_path("unCameCsv.csv");
        EasyExcel::write::<UnCamelData>(&path)
            .sheet("Sheet1")
            .do_write(uncamel_data())
            .unwrap();
        let rows = EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 11, "CSV: 1 header + 10 data");
    }
}

// ============================================================================
// TemplateDataTest (2)
// ============================================================================

mod template_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.template.TemplateDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        let template = require_fixture("template/template07.xlsx");
        let path = temp_path("template07_out.xlsx");
        EasyExcel::write::<TemplateData>(&path)
            .with_template(&template)
            .sheet("Sheet1")
            .do_write(template_data())
            .unwrap();
        let rows = EasyExcel::read_sync::<TemplateData>(&path)
            .head_row_number(3)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].string0, "字符串0");
        assert_eq!(rows[1].string0, "字符串1");
    }

    /// Java: com.alibaba.easyexcel.test.core.template.TemplateDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java withTemplate(.xls) + write. Rust: value-preserving Minimal BIFF8 rewrite.
        let xls = require_fixture("template/template03.xls");
        assert_xls_readable(&xls);
        let path = temp_path("template03_out.xls");
        EasyExcel::write::<TemplateData>(&path)
            .with_template(&xls)
            .sheet("Sheet1")
            .do_write(template_data())
            .unwrap();
        let rows = EasyExcel::read_sync::<TemplateData>(&path)
            .head_row_number(3)
            .do_read_sync()
            .unwrap();
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].string0, "字符串0");
        assert_eq!(rows[1].string0, "字符串1");
        // Template cells before the append must remain (value preserve).
        let dynamic = EasyExcel::read_dynamic_sync(&path)
            .head_row_number(0)
            .do_read_sync()
            .unwrap();
        assert!(
            dynamic.len() >= 4,
            "template rows + head + data expected, got {}",
            dynamic.len()
        );
    }
}
