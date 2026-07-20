//! Java golden cross-check — load checked-in `tests/golden/*.expected.json`
//! exported by `scripts/export-java-golden.sh` (true Java EasyExcel read/write)
//! and compare Rust read / write+read results (row_count + display cells).
//!
//! Missing golden files **fail** (no soft-skip). Run:
//! `./scripts/export-java-golden.sh` (requires JDK + Maven).
//!
//! Coverage scenarios (≥100 expected.json；ofNoRows=0):
//! - compatibility t01–t07/t09 (xlsx + xls), BOM csv, demo (xlsx/csv/extra/cellData)
//! - dataformat (xlsx/xls/v2/date1/date2), template, multi-sheet
//! - simple write (xlsx/csv/xls), converter (fixture + write xlsx/xls/csv), fill, style
//! - exclude/include, no-head(+xls/csv), sort, encrypt (password)
//! - cache / celldata(+xls/csv full) / charset / exception / handler /
//!   large-sample(+xls/csv) / nomodel / noncamel / parameter (xlsx/csv/xls) /
//!   repetition / skip / complex-head(+xls) / annotation-index(+xls) /
//!   list-head(+xls/csv) / fill-horizontal(+xls) / fill-by-name

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use easyexcel::{CellValue, DynamicRow, DynamicValue, EasyExcel, ReadDefaultReturn};
use serde::Deserialize;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

fn golden_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/golden")
}

fn golden_artifact(name: &str) -> PathBuf {
    golden_dir().join("artifacts").join(name)
}

/// JSON schema written by `scripts/java-golden-export` / `export-java-golden.sh`.
#[derive(Debug, Deserialize)]
struct GoldenExpectation {
    /// Java test class#method that owns this fixture assertion.
    source: String,
    /// Relative fixture path under `tests/fixtures` or `artifacts/...`.
    #[serde(default)]
    fixture: String,
    /// Sheet index used by the Java export.
    #[serde(default)]
    sheet_index: usize,
    /// Optional sheet name (takes precedence over `sheet_index` when set).
    #[serde(default)]
    sheet_name: Option<String>,
    /// `headRowNumber` used by the Java export.
    #[serde(default)]
    head_row_number: u32,
    /// Optional workbook password (encrypt scenarios).
    #[serde(default)]
    password: Option<String>,
    /// Optional CSV charset (e.g. `GBK` / `UTF-8`).
    #[serde(default)]
    charset: Option<String>,
    /// Number of rows returned by Java `doReadSync`.
    row_count: usize,
    /// Key cells as `"row.col" → display text` (Java STRING mode).
    #[serde(default)]
    cells: BTreeMap<String, String>,
    /// Full sheet rows as display strings (optional; compared when present).
    #[serde(default)]
    rows: Vec<Vec<String>>,
}

/// Load a golden file; **fails** if missing or invalid JSON.
fn load_golden(name: &str) -> GoldenExpectation {
    let path = golden_dir().join(name);
    assert!(
        path.is_file(),
        "required Java golden missing (run scripts/export-java-golden.sh): {}",
        path.display()
    );
    let text = std::fs::read_to_string(&path).unwrap_or_else(|e| {
        panic!("failed to read golden {}: {e}", path.display());
    });
    serde_json::from_str(&text).unwrap_or_else(|e| {
        panic!("invalid golden JSON {}: {e}", path.display());
    })
}

/// Resolve the file path referenced by a golden (`fixtures/...` or `artifacts/...`).
fn resolve_golden_path(golden: &GoldenExpectation) -> PathBuf {
    let rel = golden.fixture.as_str();
    if rel.is_empty() {
        panic!(
            "golden has empty fixture field (source={}); re-run scripts/export-java-golden.sh",
            golden.source
        );
    }
    if let Some(rest) = rel.strip_prefix("artifacts/") {
        let path = golden_artifact(rest);
        assert!(
            path.is_file(),
            "required Java write artifact missing (run scripts/export-java-golden.sh): {}",
            path.display()
        );
        return path;
    }
    let path = fixture(rel);
    assert!(
        path.is_file(),
        "required Java fixture missing: {}",
        path.display()
    );
    path
}

/// Convert a Rust `DynamicValue` to display text comparable with Java STRING mode.
fn display_text(value: &DynamicValue) -> String {
    match value {
        DynamicValue::Null => String::new(),
        DynamicValue::String(s) => s.clone(),
        DynamicValue::ActualData(cv) => cv.as_text(),
        DynamicValue::ReadCellData(cell) => cell.display_value().to_owned(),
    }
}

/// Read a path with the same sheet / head / password / charset options as the Java golden export.
fn read_display_rows(path: &Path, golden: &GoldenExpectation) -> Vec<DynamicRow> {
    let mut builder = EasyExcel::read_sync::<DynamicRow>(path)
        .head_row_number(golden.head_row_number)
        .read_default_return(ReadDefaultReturn::String);
    if let Some(password) = golden.password.as_deref() {
        if !password.is_empty() {
            builder = builder.password(password);
        }
    }
    if let Some(charset) = golden.charset.as_deref() {
        if !charset.is_empty() {
            builder = builder.charset(charset);
        }
    }
    builder = match golden.sheet_name.as_deref() {
        Some(name) if !name.is_empty() => builder.sheet(name),
        _ => builder.sheet(golden.sheet_index),
    };
    builder
        .do_read_sync()
        .unwrap_or_else(|e| panic!("Rust read failed for {}: {e}", path.display()))
}

/// Assert Rust rows match golden `row_count`, key `cells`, and full `rows` when present.
/// Date columns are compared like any other STRING cell (no soft-skip).
fn assert_matches_golden(golden: &GoldenExpectation, rows: &[DynamicRow]) {
    assert_eq!(
        rows.len(),
        golden.row_count,
        "row_count mismatch vs Java golden ({})",
        golden.source
    );

    for (coord, expected) in &golden.cells {
        let (row_idx, col_idx) = parse_coord(coord);
        let actual = rows
            .get(row_idx)
            .and_then(|r| r.get(col_idx))
            .map(display_text)
            .unwrap_or_default();
        assert_eq!(
            actual, *expected,
            "cell {coord} mismatch vs Java golden ({})",
            golden.source
        );
    }

    if golden.rows.is_empty() {
        return;
    }
    assert_eq!(
        rows.len(),
        golden.rows.len(),
        "full rows length mismatch vs Java golden ({})",
        golden.source
    );
    for (r, expected_row) in golden.rows.iter().enumerate() {
        for (c, expected_cell) in expected_row.iter().enumerate() {
            let actual = rows
                .get(r)
                .and_then(|row| row.get(c))
                .map(display_text)
                .unwrap_or_default();
            // Trailing empty columns may be omitted on the sparse side; treat missing as "".
            if actual.is_empty() && expected_cell.is_empty() {
                continue;
            }
            assert_eq!(
                actual, *expected_cell,
                "rows[{r}][{c}] mismatch vs Java golden ({})",
                golden.source
            );
        }
    }
}

/// Load golden, resolve path, read with Rust, assert full STRING match.
fn assert_golden_file(golden_name: &str) {
    let golden = load_golden(golden_name);
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Parse `"row.col"` coordinate used in golden `cells`.
fn parse_coord(coord: &str) -> (usize, usize) {
    let mut parts = coord.split('.');
    let row = parts
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or_else(|| panic!("bad golden cell coord: {coord}"));
    let col = parts
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or_else(|| panic!("bad golden cell coord: {coord}"));
    (row, col)
}

/// Sample rows matching Java `SimpleDataTest#data()` (姓名0..9).
fn simple_names() -> Vec<String> {
    (0..10).map(|i| format!("姓名{i}")).collect()
}

/// Java CompatibilityTest#t02 — fixed expected cell + full golden file.
#[test]
fn golden_compatibility_t02() {
    let path = fixture("compatibility/t02.xlsx");
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );

    let golden = load_golden("compatibility_t02.expected.json");
    assert!(
        golden.source.contains("CompatibilityTest#t02"),
        "unexpected source: {}",
        golden.source
    );

    // Hard-coded Java assertion (CompatibilityTest#t02) — keep alongside golden.
    let actual_data_rows = EasyExcel::read_sync::<DynamicRow>(&path)
        .head_row_number(0)
        .read_default_return(ReadDefaultReturn::ActualData)
        .do_read_sync()
        .unwrap();
    assert_eq!(actual_data_rows.len(), 3);
    let val = match actual_data_rows[2].get(2).unwrap() {
        DynamicValue::ActualData(CellValue::String(s)) => s.as_str(),
        DynamicValue::String(s) => s.as_str(),
        other => panic!("unexpected {other:?}"),
    };
    assert_eq!(val, "1，2-戊二醇");

    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Java CompatibilityTest#t04 — merged-cell fixture row count + key cell + golden.
#[test]
fn golden_compatibility_t04() {
    let path = fixture("compatibility/t04.xlsx");
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );

    let golden = load_golden("compatibility_t04.expected.json");
    assert!(
        golden.source.contains("CompatibilityTest#t04"),
        "unexpected source: {}",
        golden.source
    );

    let actual_data_rows = EasyExcel::read_sync::<DynamicRow>(&path)
        .read_default_return(ReadDefaultReturn::ActualData)
        .do_read_sync()
        .unwrap();
    assert_eq!(actual_data_rows.len(), 56);
    let val = match actual_data_rows[0].get(5).unwrap() {
        DynamicValue::ActualData(CellValue::String(s)) => s.as_str(),
        DynamicValue::String(s) => s.as_str(),
        other => panic!("unexpected {other:?}"),
    };
    assert_eq!(val, "QQSJK28F152A012242S0081");

    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CompatibilityTest#t01 — Java .xls fixture STRING read对照.
#[test]
fn golden_compatibility_t01_xls() {
    let golden = load_golden("compatibility_t01_xls.expected.json");
    assert!(
        golden.source.contains("CompatibilityTest#t01"),
        "unexpected source: {}",
        golden.source
    );
    assert!(
        golden.fixture.ends_with(".xls"),
        "expected .xls fixture, got {}",
        golden.fixture
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// BOM fixtures must match Java BomDataTest expectations + Java golden JSON.
#[test]
fn golden_bom_office_csv() {
    #[derive(Debug, Clone, easyexcel::ExcelRow)]
    struct BomData {
        #[excel(name = "姓名")]
        name: String,
        #[excel(name = "年纪")]
        age: i64,
    }
    let path = fixture("bom/office_bom.csv");
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );

    let typed = EasyExcel::read_sync::<BomData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(typed.len(), 10);
    assert_eq!(typed[0].name, "姓名0");
    assert_eq!(typed[0].age, 20);

    let golden = load_golden("bom_office_bom.expected.json");
    assert!(
        golden.source.contains("BomDataTest"),
        "unexpected source: {}",
        golden.source
    );
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// no_bom.csv — same logical content as office_bom without UTF-8 BOM.
#[test]
fn golden_bom_no_bom_csv() {
    assert_golden_file("bom_no_bom.expected.json");
}

/// Java ReadTest#simpleRead — demo.xlsx first sheet vs Java golden (date col included).
#[test]
fn golden_demo_demo_sheet0() {
    let golden = load_golden("demo_demo_sheet0.expected.json");
    assert!(
        golden
            .source
            .contains("com.alibaba.easyexcel.test.demo.read.ReadTest"),
        "unexpected source: {}",
        golden.source
    );
    assert!(
        golden.cells.contains_key("1.1"),
        "demo golden must include date cell 1.1 for STRING alignment"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// demo.csv — full STRING-mode display vs Java golden including date column.
#[test]
fn golden_demo_demo_csv() {
    let golden = load_golden("demo_demo_csv.expected.json");
    assert!(
        golden.cells.contains_key("1.1"),
        "demo csv golden must include date cell 1.1"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// SimpleDataTest#t21 — sheet name `simple` on simple07.xlsx.
#[test]
fn golden_simple_simple07_sheet_name() {
    assert_golden_file("simple_simple07.expected.json");
}

/// Java-written SimpleData xlsx artifact must match golden; Rust write+read must match too.
#[test]
fn golden_simple_data_xlsx_write() {
    #[derive(Debug, Clone, easyexcel::ExcelRow)]
    struct SimpleData {
        #[excel(name = "姓名")]
        name: String,
    }

    let golden = load_golden("simple_data.expected.json");
    assert!(
        golden.source.contains("SimpleDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert_eq!(golden.row_count, 10);

    let java_artifact = golden_artifact("simple_data.xlsx");
    assert!(
        java_artifact.is_file(),
        "required Java write artifact missing (run scripts/export-java-golden.sh): {}",
        java_artifact.display()
    );
    let java_rows = read_display_rows(&java_artifact, &golden);
    assert_matches_golden(&golden, &java_rows);

    let path = tempfile::tempdir()
        .unwrap()
        .keep()
        .join("simple_golden.xlsx");
    let data: Vec<SimpleData> = simple_names()
        .into_iter()
        .map(|name| SimpleData { name })
        .collect();
    EasyExcel::write::<SimpleData>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rust_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &rust_rows);
}

/// Java-written SimpleData csv artifact + Rust csv write/read vs golden.
#[test]
fn golden_simple_data_csv_write() {
    #[derive(Debug, Clone, easyexcel::ExcelRow)]
    struct SimpleData {
        #[excel(name = "姓名")]
        name: String,
    }

    let golden = load_golden("simple_data_csv.expected.json");
    assert!(
        golden.source.contains("SimpleDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert_eq!(golden.row_count, 10);

    let java_artifact = golden_artifact("simple_data.csv");
    assert!(
        java_artifact.is_file(),
        "required Java write artifact missing (run scripts/export-java-golden.sh): {}",
        java_artifact.display()
    );
    let java_rows = read_display_rows(&java_artifact, &golden);
    assert_matches_golden(&golden, &java_rows);

    let path = tempfile::tempdir()
        .unwrap()
        .keep()
        .join("simple_golden.csv");
    let data: Vec<SimpleData> = simple_names()
        .into_iter()
        .map(|name| SimpleData { name })
        .collect();
    EasyExcel::write::<SimpleData>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rust_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &rust_rows);
}

/// Java-written SimpleData `.xls` artifact — Rust **read**对照 only (write unsupported).
#[test]
fn golden_simple_data_xls_read() {
    let golden = load_golden("simple_data_xls.expected.json");
    assert!(
        golden.source.contains("SimpleDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert!(
        golden.fixture.ends_with(".xls"),
        "expected .xls artifact, got {}",
        golden.fixture
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Converter fixture converter07.xlsx — STRING read including date columns.
#[test]
fn golden_converter_converter07() {
    let golden = load_golden("converter_converter07.expected.json");
    assert!(
        golden.source.contains("ConverterDataTest"),
        "unexpected source: {}",
        golden.source
    );
    // Date columns (e.g. 0.12 / 0.13) must be present and compared.
    assert!(
        golden.cells.contains_key("0.12") && golden.cells.contains_key("0.13"),
        "converter golden must include date STRING cells"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Converter `.xls` fixture — full-table STRING including short dates (`xls_display`).
#[test]
fn golden_converter_converter03_xls() {
    let golden = load_golden("converter_converter03_xls.expected.json");
    assert!(
        golden.fixture.ends_with(".xls"),
        "expected .xls fixture, got {}",
        golden.fixture
    );
    assert_eq!(
        golden.cells.get("0.12").map(String::as_str),
        Some("2020-1-1 1:01")
    );
    assert_eq!(
        golden.cells.get("0.13").map(String::as_str),
        Some("2020-01-01 01:01:01")
    );
    assert!(
        !golden.rows.is_empty(),
        "converter03.xls must be full-table STRING"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Converter csv fixture.
#[test]
fn golden_converter_converter_csv() {
    assert_golden_file("converter_converter_csv.expected.json");
}

/// Java ConverterWriteData artifact — date / localDate / localDateTime STRING对齐.
#[test]
fn golden_converter_write() {
    let golden = load_golden("converter_write.expected.json");
    assert!(
        golden.cells.get("0.0").map(String::as_str) == Some("2020-01-01 01:01:01"),
        "converter write golden date cell missing/mismatched"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// FillTest#simpleFill — Java filled artifact vs Rust STRING read.
#[test]
fn golden_fill_simple() {
    let golden = load_golden("fill_simple.expected.json");
    assert!(
        golden.source.contains("FillTest"),
        "unexpected source: {}",
        golden.source
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Multi-sheet xlsx — sheet 0.
#[test]
fn golden_multiplesheets_sheet0() {
    assert_golden_file("multiplesheets_sheet0.expected.json");
}

/// Multi-sheet xlsx — sheet 1.
#[test]
fn golden_multiplesheets_sheet1() {
    assert_golden_file("multiplesheets_sheet1.expected.json");
}

/// Multi-sheet `.xls` — sheet 0 Rust read.
#[test]
fn golden_multiplesheets_xls_sheet0() {
    assert_golden_file("multiplesheets_xls_sheet0.expected.json");
}

/// Multi-sheet `.xls` — sheet 1 Rust read.
#[test]
fn golden_multiplesheets_xls_sheet1() {
    assert_golden_file("multiplesheets_xls_sheet1.expected.json");
}

/// CompatibilityTest#t03 — sparse null leading columns.
#[test]
fn golden_compatibility_t03() {
    assert_golden_file("compatibility_t03.expected.json");
}

/// CompatibilityTest#t05 — date rounding full STRING cells.
#[test]
fn golden_compatibility_t05_dates() {
    let golden = load_golden("compatibility_t05.expected.json");
    assert!(
        golden.cells.get("0.0").map(String::as_str) == Some("2023-01-01 00:00:00"),
        "t05 date STRING cell missing"
    );
    assert!(
        golden.cells.get("3.0").map(String::as_str) == Some("2023-01-01 00:00:01"),
        "t05 rounded second cell missing"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CompatibilityTest#t06 — numeric precision STRING.
#[test]
fn golden_compatibility_t06() {
    let golden = load_golden("compatibility_t06.expected.json");
    assert_eq!(
        golden.cells.get("0.2").map(String::as_str),
        Some("2087.03")
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CompatibilityTest#t07 — STRING "24.20" + trailing-space accounting (`-1.07 `).
#[test]
fn golden_compatibility_t07() {
    let golden = load_golden("compatibility_t07.expected.json");
    assert_eq!(golden.cells.get("0.11").map(String::as_str), Some("24.20"));
    assert_eq!(golden.cells.get("0.12").map(String::as_str), Some("-1.07 "));
    assert!(!golden.rows.is_empty(), "t07 must be full-table STRING");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CompatibilityTest#t09 — sharedStrings escape.
#[test]
fn golden_compatibility_t09() {
    let golden = load_golden("compatibility_t09.expected.json");
    assert_eq!(
        golden.cells.get("0.0").map(String::as_str),
        Some("SH_x000D_Z002")
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// DateFormatTest#t03Read — full STRING including unpadded month `2023-1-01`.
#[test]
fn golden_dataformat_v2() {
    let golden = load_golden("dataformat_v2.expected.json");
    assert_eq!(golden.cells.get("0.0").map(String::as_str), Some("15:00"));
    assert_eq!(
        golden.cells.get("1.0").map(String::as_str),
        Some("2023-1-01 00:00:00")
    );
    assert!(!golden.rows.is_empty(), "dataformatv2 must be full-table STRING");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// DateFormatTest#t01 — dataformat.xlsx full STRING (CN AM/PM, mmmmm PUA, ￥).
#[test]
fn golden_dataformat_xlsx() {
    let golden = load_golden("dataformat_xlsx.expected.json");
    assert_eq!(golden.cells.get("22.0").map(String::as_str), Some("上午1时01分"));
    assert!(!golden.rows.is_empty(), "dataformat_xlsx must be full-table STRING");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// DateFormatTest#t02 — dataformat.xls full STRING (BIFF ¥, CN AM/PM, mmmmm PUA).
#[test]
fn golden_dataformat_xls() {
    let golden = load_golden("dataformat_xls.expected.json");
    assert_eq!(golden.cells.get("2.4").map(String::as_str), Some("¥1.11"));
    assert_eq!(golden.cells.get("22.0").map(String::as_str), Some("上午1时01分"));
    assert!(!golden.rows.is_empty(), "dataformat_xls must be full-table STRING");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// issue2443 date1.xlsx STRING.
#[test]
fn golden_dataformat_date1() {
    assert_golden_file("dataformat_date1.expected.json");
}

/// issue2443 date2.xlsx STRING.
#[test]
fn golden_dataformat_date2() {
    assert_golden_file("dataformat_date2.expected.json");
}

/// ExtraDataTest content (xlsx).
#[test]
fn golden_demo_extra_xlsx() {
    assert_golden_file("demo_extra_xlsx.expected.json");
}

/// ExtraDataTest content (xls).
#[test]
fn golden_demo_extra_xls() {
    assert_golden_file("demo_extra_xls.expected.json");
}

/// cellDataDemo.xlsx.
#[test]
fn golden_demo_cell_data() {
    assert_golden_file("demo_cell_data.expected.json");
}

/// demo/simple07.xlsx sheet `simple`.
#[test]
fn golden_demo_simple07() {
    assert_golden_file("demo_simple07.expected.json");
}

/// template07.xlsx content read.
#[test]
fn golden_template_template07() {
    assert_golden_file("template_template07.expected.json");
}

/// template03.xls content read.
#[test]
fn golden_template_template03_xls() {
    assert_golden_file("template_template03_xls.expected.json");
}

/// StyleDataTest write artifact — STRING content对照 (styles are write-side).
#[test]
fn golden_style_data() {
    let golden = load_golden("style_data.expected.json");
    assert!(
        golden.source.contains("StyleDataTest"),
        "unexpected source: {}",
        golden.source
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// StyleDataTest `.xls` artifact — Rust read对照.
#[test]
fn golden_style_data_xls() {
    assert_golden_file("style_data_xls.expected.json");
}

/// AnnotationData — DateTimeFormat + `#.##%` full STRING (`java_compat_display` → `9999%`).
#[test]
fn golden_annotation_data() {
    let golden = load_golden("annotation_data.expected.json");
    assert!(
        golden.source.contains("AnnotationDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert!(
        golden.cells.get("0.0").is_some_and(|s| s.contains("年")),
        "annotation date format cell missing"
    );
    assert_eq!(golden.cells.get("0.1").map(String::as_str), Some("9999%"));
    assert!(
        !golden.rows.is_empty(),
        "annotation_data must be full-table STRING"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
    // Confirm ignore column was not written (only date + number).
    assert!(
        display_rows[0]
            .get(2)
            .map(display_text)
            .unwrap_or_default()
            .is_empty(),
        "ExcelIgnore field must not appear as a third column"
    );
}

/// ExcludeOrInclude excludeColumnIndexes — only column2/column3 remain.
#[test]
fn golden_exclude_index() {
    let golden = load_golden("exclude_index.expected.json");
    assert!(
        golden.source.contains("ExcludeOrIncludeDataTest"),
        "unexpected source: {}",
        golden.source
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// ExcludeOrInclude exclude index CSV.
#[test]
fn golden_exclude_index_csv() {
    assert_golden_file("exclude_index_csv.expected.json");
}

/// ExcludeOrInclude exclude field names.
#[test]
fn golden_exclude_field() {
    assert_golden_file("exclude_field.expected.json");
}

/// ExcludeOrInclude include indexes.
#[test]
fn golden_include_index() {
    assert_golden_file("include_index.expected.json");
}

/// ExcludeOrInclude include field names.
#[test]
fn golden_include_field() {
    assert_golden_file("include_field.expected.json");
}

/// ExcludeOrInclude include field names with order.
#[test]
fn golden_include_field_order() {
    assert_golden_file("include_field_order.expected.json");
}

/// FillDataTest#t02Fill03 — Java filled `.xls` artifact.
#[test]
fn golden_fill_simple_xls() {
    assert_golden_file("fill_simple_xls.expected.json");
}

/// FillDataTest#t05HorizontalFill07.
#[test]
fn golden_fill_horizontal() {
    let golden = load_golden("fill_horizontal.expected.json");
    assert_eq!(golden.cells.get("0.2").map(String::as_str), Some("张三"));
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// NoHeadDataTest — needHead(false).
#[test]
fn golden_no_head_data() {
    assert_golden_file("no_head_data.expected.json");
}

/// SortDataTest — index/order columns.
#[test]
fn golden_sort_data() {
    assert_golden_file("sort_data.expected.json");
}

/// EncryptDataTest — Java encrypted artifact; Rust read with password.
#[test]
fn golden_encrypt_data() {
    let golden = load_golden("encrypt_data.expected.json");
    assert!(
        golden.source.contains("EncryptDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert_eq!(
        golden.password.as_deref(),
        Some("123456"),
        "encrypt golden must carry password"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CacheDataTest — 姓名/年龄 full-table STRING.
#[test]
fn golden_cache_data() {
    let golden = load_golden("cache_data.expected.json");
    assert!(
        golden.source.contains("CacheDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert!(!golden.rows.is_empty(), "cache golden must include full rows");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CacheDataTest `.xls` / `.csv` 格式变体.
#[test]
fn golden_cache_data_xls_csv() {
    assert_golden_file("cache_data_xls.expected.json");
    assert_golden_file("cache_data_csv.expected.json");
}

/// CellDataDataTest xlsx — date/number/formula STRING full table.
#[test]
fn golden_celldata_data() {
    let golden = load_golden("celldata_data.expected.json");
    assert!(
        golden.source.contains("CellDataDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert!(!golden.rows.is_empty(), "celldata golden must include full rows");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CharsetDataTest GBK CSV — charset field required.
#[test]
fn golden_charset_gbk() {
    let golden = load_golden("charset_gbk.expected.json");
    assert_eq!(golden.charset.as_deref(), Some("GBK"));
    assert!(!golden.rows.is_empty(), "charset golden must include full rows");
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CharsetDataTest UTF-8 CSV.
#[test]
fn golden_charset_utf8() {
    let golden = load_golden("charset_utf8.expected.json");
    assert_eq!(golden.charset.as_deref(), Some("UTF-8"));
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// ExceptionDataTest content.
#[test]
fn golden_exception_data() {
    assert_golden_file("exception_data.expected.json");
}

/// ExceptionDataTest multi-sheet stop fixture — sheet0.
#[test]
fn golden_exception_stop_sheet0() {
    let golden = load_golden("exception_stop_sheet0.expected.json");
    assert!(
        golden.source.contains("ExceptionDataTest"),
        "unexpected source: {}",
        golden.source
    );
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// WriteHandlerTest content — 姓名0..9 full table.
#[test]
fn golden_handler_data() {
    let golden = load_golden("handler_data.expected.json");
    assert_eq!(golden.row_count, 10);
    assert!(!golden.rows.is_empty());
    assert_eq!(golden.cells.get("9.0").map(String::as_str), Some("姓名9"));
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// WriteHandlerTest CSV content.
#[test]
fn golden_handler_data_csv() {
    assert_golden_file("handler_data_csv.expected.json");
}

/// LargeDataTest sample (100×25) — not large07.
#[test]
fn golden_large_sample() {
    let golden = load_golden("large_sample.expected.json");
    assert_eq!(golden.row_count, 100);
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// LargeDataTest CSV sample (100×25).
#[test]
fn golden_large_sample_csv() {
    let golden = load_golden("large_sample_csv.expected.json");
    assert_eq!(golden.row_count, 100);
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Converter write `.xls` artifact — full STRING.
#[test]
fn golden_converter_write_xls() {
    assert_golden_file("converter_write_xls.expected.json");
}

/// Converter write CSV artifact — full STRING.
#[test]
fn golden_converter_write_csv() {
    assert_golden_file("converter_write_csv.expected.json");
}

/// CellDataDataTest `.xls` — full STRING（CN DateTimeFormat 已对齐）.
#[test]
fn golden_celldata_data_xls() {
    let golden = load_golden("celldata_data_xls.expected.json");
    assert!(!golden.rows.is_empty(), "celldata_xls must export full rows");
    assert!(
        golden.cells.get("0.0").is_some_and(|s| s.contains("年")),
        "celldata xls must keep CN date text"
    );
    assert_eq!(golden.cells.get("0.1").map(String::as_str), Some("2"));
    assert_eq!(golden.cells.get("0.2").map(String::as_str), Some("2"));
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// CellDataDataTest CSV — full STRING (literal CN date text).
#[test]
fn golden_celldata_data_csv() {
    let golden = load_golden("celldata_data_csv.expected.json");
    assert!(!golden.rows.is_empty());
    assert!(
        golden.cells.get("0.0").is_some_and(|s| s.contains("年")),
        "celldata csv must keep CN date text"
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// ComplexHeadDataTest — multi-level head, headRowNumber(3).
#[test]
fn golden_complex_head() {
    let golden = load_golden("complex_head.expected.json");
    assert_eq!(golden.head_row_number, 3);
    assert!(!golden.rows.is_empty());
    assert_eq!(golden.cells.get("0.4").map(String::as_str), Some("字符串4"));
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// AnnotationIndexAndNameDataTest — sparse column index 4.
#[test]
fn golden_annotation_index_name() {
    let golden = load_golden("annotation_index_name.expected.json");
    assert_eq!(golden.cells.get("0.0").map(String::as_str), Some("第0个"));
    assert_eq!(golden.cells.get("0.4").map(String::as_str), Some("第4个"));
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
    // Column 3 is intentionally empty (no @ExcelProperty index=3).
    assert!(
        display_rows[0]
            .get(3)
            .map(display_text)
            .unwrap_or_default()
            .is_empty(),
        "index gap at col3 must be empty"
    );
}

/// ListHeadDataTest xlsx — full STRING including date + 额外数据.
#[test]
fn golden_list_head() {
    let golden = load_golden("list_head.expected.json");
    assert!(!golden.rows.is_empty(), "list_head xlsx must export full rows");
    assert_eq!(golden.cells.get("0.0").map(String::as_str), Some("字符串0"));
    assert_eq!(
        golden.cells.get("0.2").map(String::as_str),
        Some("2020-01-01 01:01:01")
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// ListHeadDataTest `.xls` — full STRING.
#[test]
fn golden_list_head_xls() {
    assert_golden_file("list_head_xls.expected.json");
}

/// ComplexHeadDataTest `.xls` — headRowNumber(3).
#[test]
fn golden_complex_head_xls() {
    let golden = load_golden("complex_head_xls.expected.json");
    assert_eq!(golden.head_row_number, 3);
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// AnnotationIndexAndNameDataTest `.xls`.
#[test]
fn golden_annotation_index_name_xls() {
    assert_golden_file("annotation_index_name_xls.expected.json");
}

/// LargeDataTest sample `.xls` (100×25).
#[test]
fn golden_large_sample_xls() {
    let golden = load_golden("large_sample_xls.expected.json");
    assert_eq!(golden.row_count, 100);
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// NoHeadDataTest `.xls` / CSV.
#[test]
fn golden_no_head_data_xls() {
    assert_golden_file("no_head_data_xls.expected.json");
}

/// NoHeadDataTest CSV.
#[test]
fn golden_no_head_data_csv() {
    assert_golden_file("no_head_data_csv.expected.json");
}

/// FillDataTest horizontal `.xls`.
#[test]
fn golden_fill_horizontal_xls() {
    assert_golden_file("fill_horizontal_xls.expected.json");
}

/// FillDataTest byName Sheet2.
#[test]
fn golden_fill_by_name() {
    let golden = load_golden("fill_by_name.expected.json");
    assert_eq!(golden.sheet_name.as_deref(), Some("Sheet2"));
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// FillDataTest#t08ByNameFill03 + #t03/#t04 complex fill.
#[test]
fn golden_fill_by_name_xls_and_complex() {
    assert_golden_file("fill_by_name_xls.expected.json");
    let complex = load_golden("fill_complex.expected.json");
    assert_eq!(complex.head_row_number, 3);
    assert!(!complex.rows.is_empty());
    assert_golden_file("fill_complex.expected.json");
    assert_golden_file("fill_complex_xls.expected.json");
}

/// ListHeadDataTest CSV — full STRING including date text.
#[test]
fn golden_list_head_csv() {
    let golden = load_golden("list_head_csv.expected.json");
    assert!(!golden.rows.is_empty());
    assert_eq!(
        golden.cells.get("0.2").map(String::as_str),
        Some("2020-01-01 01:01:01")
    );
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// ParameterDataTest `.xls` read对照.
#[test]
fn golden_parameter_data_xls() {
    assert_golden_file("parameter_data_xls.expected.json");
}

/// NoModelDataTest — headRowNumber(0) full table.
#[test]
fn golden_nomodel_data() {
    let golden = load_golden("nomodel_data.expected.json");
    assert_eq!(golden.head_row_number, 0);
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// NoModelDataTest repeat 写回 — xlsx/xls/csv.
#[test]
fn golden_nomodel_repeat_variants() {
    assert_golden_file("nomodel_repeat.expected.json");
    assert_golden_file("nomodel_repeat_xls.expected.json");
    assert_golden_file("nomodel_repeat_csv.expected.json");
}

/// UnCamelDataTest.
#[test]
fn golden_noncamel_data() {
    assert_golden_file("noncamel_data.expected.json");
}

/// ParameterDataTest.
#[test]
fn golden_parameter_data() {
    assert_golden_file("parameter_data.expected.json");
}

/// RepetitionDataTest — double write → 2 data rows.
#[test]
fn golden_repetition_data() {
    let golden = load_golden("repetition_data.expected.json");
    assert_eq!(golden.row_count, 2);
    assert!(!golden.rows.is_empty());
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// SkipDataTest — sheet name `第二个`.
#[test]
fn golden_skip_sheet1() {
    let golden = load_golden("skip_sheet1.expected.json");
    assert_eq!(golden.sheet_name.as_deref(), Some("第二个"));
    assert_eq!(golden.cells.get("0.0").map(String::as_str), Some("name2"));
    let path = resolve_golden_path(&golden);
    let display_rows = read_display_rows(&path, &golden);
    assert_matches_golden(&golden, &display_rows);
}

/// Every checked-in `*.expected.json` must pass (guards coverage ≥100, no soft-skip).
#[test]
fn golden_all_expected_json_files() {
    let dir = golden_dir();
    let mut names: Vec<String> = std::fs::read_dir(&dir)
        .unwrap_or_else(|e| panic!("golden dir missing {}: {e}", dir.display()))
        .filter_map(|e| e.ok())
        .filter(|e| {
            e.path()
                .extension()
                .and_then(|x| x.to_str())
                .is_some_and(|ext| ext.eq_ignore_ascii_case("json"))
                && e.file_name()
                    .to_string_lossy()
                    .ends_with(".expected.json")
        })
        .map(|e| e.file_name().to_string_lossy().into_owned())
        .collect();
    names.sort();
    assert!(
        names.len() >= 108,
        "expected ≥108 golden JSON files, found {} — run scripts/export-java-golden.sh",
        names.len()
    );
    for name in &names {
        assert_golden_file(name);
    }
}

/// Missing golden must fail loudly (regression guard for soft-skip).
#[test]
#[should_panic(expected = "required Java golden missing")]
fn golden_missing_file_fails() {
    let _ = load_golden("__does_not_exist__.expected.json");
}


/// P0 STRING 全表回归：dataformat / annotation / converter03 / t07 / celldata / list_head。
#[test]
fn golden_p0_format_full_rows() {
    for (label, name) in [
        ("dataformat_v2", "dataformat_v2.expected.json"),
        ("dataformat_xlsx", "dataformat_xlsx.expected.json"),
        ("dataformat_xls", "dataformat_xls.expected.json"),
        ("annotation_data", "annotation_data.expected.json"),
        ("converter03_xls", "converter_converter03_xls.expected.json"),
        ("compatibility_t07", "compatibility_t07.expected.json"),
        ("celldata_csv", "celldata_data_csv.expected.json"),
        ("celldata_xls", "celldata_data_xls.expected.json"),
        ("list_head", "list_head.expected.json"),
        ("list_head_csv", "list_head_csv.expected.json"),
        ("list_head_xls", "list_head_xls.expected.json"),
    ] {
        let golden = load_golden(name);
        assert!(
            !golden.rows.is_empty(),
            "{label} must export full rows (ofNoRows cleared)"
        );
        let path = resolve_golden_path(&golden);
        let rows = read_display_rows(&path, &golden);
        assert_matches_golden(&golden, &rows);
    }
}
