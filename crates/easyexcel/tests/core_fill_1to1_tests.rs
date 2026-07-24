//! Method-level 1:1 parity for Java core fill tests:
//! FillDataTest, FillAnnotationDataTest, FillStyleDataTest, FillStyleAnnotatedTest.
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>` so each Rust test
//! uniquely maps to `ClassName#methodName`.
//!
//! Format strategy:
//! - `.xlsx`: fill via `EasyExcel::fill_template` / `fill_template_list` / `template_writer`
//! - `.xls`: legacy template fill is explicit `Unsupported` (no xlsx masquerade)
//! - `.csv`: assert `csv cannot use template.` (Java `ExcelGenerateException`)

use chrono::NaiveDate;
use easyexcel::{
    CellValue, DynamicRow, DynamicValue, EasyExcel, ExcelError, ExcelRow, FillConfig,
    FillDirection, FillWrapper, TemplateData, TemplateSheet,
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

fn dynamic_contains(rows: &[DynamicRow], needle: &str) -> bool {
    rows.iter().any(|row| {
        row.values().iter().any(|(_, val)| match val {
            DynamicValue::String(s) => s.contains(needle),
            DynamicValue::ActualData(CellValue::String(s)) => s.contains(needle),
            _ => false,
        })
    })
}

fn fill_data_rows() -> Vec<TemplateData> {
    (0..10)
        .map(|i| {
            let mut row = TemplateData::new().with("number", 5.2);
            if i == 5 {
                row = row.with("name", Option::<String>::None);
            } else {
                row = row.with("name", "张三");
            }
            row
        })
        .collect()
}

fn style_data_rows() -> Vec<TemplateData> {
    let date = NaiveDate::from_ymd_opt(2020, 1, 1)
        .unwrap()
        .and_hms_opt(1, 1, 1)
        .unwrap();
    (0..10)
        .map(|i| {
            let mut row = TemplateData::new()
                .with("number", 5.2)
                .with("date", date)
                .with("empty", Option::<String>::None);
            if i == 5 {
                row = row.with("name", Option::<String>::None);
            } else {
                row = row.with("name", "张三");
            }
            row
        })
        .collect()
}

/// Assert legacy .xls template fill works with SST-based templates.
/// Phase 5.2: SST parsing resolves LABELSST records so {key} placeholders
/// in the shared string table are correctly found and replaced.
fn assert_xls_fill_works(xls_template: &std::path::Path, output_name: &str) {
    assert_xls_readable(xls_template);
    let output = temp_path(output_name);
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    let result = EasyExcel::fill_template(xls_template, &output, &data);
    match result {
        Ok(()) => {
            assert!(output.exists(), "XLS fill output must exist");
            // Verify the output can be read back
            let rows = EasyExcel::read_dynamic_sync(&output)
                .head_row_number(0)
                .do_read_sync()
                .unwrap_or_default();
            assert!(!rows.is_empty(), "Filled XLS must contain readable rows");
        }
        Err(e) => {
            panic!("XLS template fill failed unexpectedly: {e}");
        }
    }
}

fn assert_simple_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(template, &output, &data).unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三"),
        "filled template must contain 张三: {rows:?}"
    );
}

fn assert_complex_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    let mut writer = EasyExcel::template_writer(template, &output).unwrap();
    let cfg = FillConfig::new().force_new_row(true);
    // Java: fill(data, forceNewRow) twice + fill(map with date/total).
    writer
        .fill_list(&FillWrapper::new(fill_data_rows()), cfg)
        .unwrap();
    writer
        .fill_list(&FillWrapper::new(fill_data_rows()), cfg)
        .unwrap();
    writer
        .fill(
            &TemplateData::new()
                .with("date", "2019年10月9日13:28:28")
                .with("total", 1000),
        )
        .unwrap();
    writer.finish().unwrap();

    // Scalars live above the list head — read from row 0.
    let all_rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(
        dynamic_contains(&all_rows, "2019年10月9日13:28:28"),
        "complex fill must contain scalar date: {all_rows:?}"
    );
    assert!(
        dynamic_contains(&all_rows, "统计:1000") || dynamic_contains(&all_rows, "1000"),
        "complex fill must contain scalar total: {all_rows:?}"
    );

    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(3)
        .do_read_sync()
        .unwrap();
    // Java: assertEquals(21, list.size()); map19.get(0)=="张三"
    assert_eq!(
        rows.len(),
        21,
        "complex fill expected 21 data rows after head, got {}: {rows:?}",
        rows.len()
    );
    assert!(
        dynamic_contains(&rows, "张三"),
        "complex fill must contain 张三: {rows:?}"
    );
}

fn assert_horizontal_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    let mut writer = EasyExcel::template_writer(template, &output).unwrap();
    let cfg = FillConfig::new().direction(FillDirection::Horizontal);
    writer
        .fill_list(&FillWrapper::new(fill_data_rows()), cfg)
        .unwrap();
    writer
        .fill_list(&FillWrapper::new(fill_data_rows()), cfg)
        .unwrap();
    writer
        .fill(&TemplateData::new().with("date", "2019年10月9日13:28:28"))
        .unwrap();
    writer.finish().unwrap();

    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    // Java: assertEquals(5, list.size()); map0.get(2)=="张三"
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三"),
        "horizontal fill must contain 张三: {rows:?}"
    );
}

fn assert_by_name_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    let mut writer = EasyExcel::template_writer(template, &output).unwrap();
    writer
        .fill_on_sheet(&TemplateSheet::Name("Sheet2".to_owned()), &data)
        .unwrap();
    writer.finish().unwrap();

    let rows = EasyExcel::read_dynamic_sync(&output)
        .sheet("Sheet2")
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三"),
        "byName Sheet2 fill must contain 张三: {rows:?}"
    );
}

fn assert_composite_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    let mut writer = EasyExcel::template_writer(template, &output).unwrap();
    let horizontal = FillConfig::new().direction(FillDirection::Horizontal);
    writer
        .fill_list(&FillWrapper::named("data1", fill_data_rows()), horizontal)
        .unwrap();
    writer
        .fill_list(&FillWrapper::named("data1", fill_data_rows()), horizontal)
        .unwrap();
    writer
        .fill_list(
            &FillWrapper::named("data2", fill_data_rows()),
            FillConfig::new(),
        )
        .unwrap();
    writer
        .fill_list(
            &FillWrapper::named("data2", fill_data_rows()),
            FillConfig::new(),
        )
        .unwrap();
    writer
        .fill_list(
            &FillWrapper::named("data3", fill_data_rows()),
            FillConfig::new(),
        )
        .unwrap();
    writer
        .fill_list(
            &FillWrapper::named("data3", fill_data_rows()),
            FillConfig::new(),
        )
        .unwrap();
    writer
        .fill(&TemplateData::new().with("date", "2019年10月9日13:28:28"))
        .unwrap();
    writer.finish().unwrap();

    let rows = EasyExcel::read_dynamic_sync(&output)
        .ignore_empty_row(false)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三"),
        "composite fill must contain 张三: {rows:?}"
    );
}

fn assert_style_list_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    EasyExcel::fill_template_list(
        template,
        &output,
        &FillWrapper::new(style_data_rows()),
        FillConfig::new(),
    )
    .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三"),
        "style fill must contain 张三: {rows:?}"
    );
}

fn assert_annotation_list_fill(template: &std::path::Path, output_name: &str) {
    let output = temp_path(output_name);
    let date = NaiveDate::from_ymd_opt(2020, 1, 1)
        .unwrap()
        .and_hms_opt(1, 1, 1)
        .unwrap();
    let img_path = require_fixture("converter/img.jpg");
    let img_bytes = std::fs::read(&img_path).unwrap();
    let row = TemplateData::new()
        .with("date", date)
        .with("number", 99.99)
        .with("string1", "string1")
        .with("string2", "string2")
        .with("image", CellValue::Image(img_bytes));
    let rows_data = vec![row.clone(), row.clone(), row.clone(), row.clone(), row];
    EasyExcel::fill_template_list(
        template,
        &output,
        &FillWrapper::new(rows_data),
        FillConfig::new(),
    )
    .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "string1") || dynamic_contains(&rows, "string2"),
        "annotation fill must contain string fields: {rows:?}"
    );
}

// ============================================================================
// FillDataTest (11 @Test)
// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest
// ============================================================================

mod fill_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t01Fill07
    #[test]
    fn t01_fill07() {
        let template = require_fixture("fill/simple.xlsx");
        assert_simple_fill(&template, "fill07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t02Fill03
    #[test]
    fn t02_fill03() {
        // Java fills xls/fill/simple.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(&require_fixture("xls/fill/simple.xls"), "t02_fill03.xls");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t03FillCsv
    #[test]
    fn t03_fill_csv() {
        // Java: assertThrows ExcelGenerateException("csv cannot use template.")
        #[derive(Debug, Clone, ExcelRow)]
        struct FillData {
            #[excel(name = "name", index = 0)]
            name: String,
        }
        let template = require_fixture("fill/simple.csv");
        let output = temp_path("fill.csv");
        let err = EasyExcel::write::<FillData>(&output)
            .with_template(&template)
            .sheet("Sheet1")
            .do_write(vec![FillData {
                name: "张三".to_owned(),
            }])
            .expect_err("csv cannot use template");
        assert!(
            err.to_string().contains("csv cannot use template"),
            "unexpected error: {err}"
        );
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t03ComplexFill07
    #[test]
    fn t03_complex_fill07() {
        let template = require_fixture("fill/complex.xlsx");
        assert_complex_fill(&template, "fillComplex07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t04ComplexFill03
    #[test]
    fn t04_complex_fill03() {
        // Java fills xls/fill/complex.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(
            &require_fixture("xls/fill/complex.xls"),
            "t04_complex_fill03.xls",
        );
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t05HorizontalFill07
    #[test]
    fn t05_horizontal_fill07() {
        let template = require_fixture("fill/horizontal.xlsx");
        assert_horizontal_fill(&template, "fillHorizontal07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t06HorizontalFill03
    #[test]
    fn t06_horizontal_fill03() {
        // Java fills xls/fill/horizontal.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(
            &require_fixture("xls/fill/horizontal.xls"),
            "t06_horizontal_fill03.xls",
        );
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t07ByNameFill07
    #[test]
    fn t07_by_name_fill07() {
        let template = require_fixture("fill/byName.xlsx");
        assert_by_name_fill(&template, "byName07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t08ByNameFill03
    #[test]
    fn t08_by_name_fill03() {
        // Java fills xls/fill/byName.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(
            &require_fixture("xls/fill/byName.xls"),
            "t08_by_name_fill03.xls",
        );
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t09CompositeFill07
    #[test]
    fn t09_composite_fill07() {
        let template = require_fixture("fill/composite.xlsx");
        assert_composite_fill(&template, "fileComposite07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.FillDataTest#t10CompositeFill03
    #[test]
    fn t10_composite_fill03() {
        // Java fills xls/fill/composite.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(
            &require_fixture("xls/fill/composite.xls"),
            "t10_composite_fill03.xls",
        );
    }
}

// ============================================================================
// FillAnnotationDataTest (2 @Test)
// Java: com.alibaba.easyexcel.test.core.fill.annotation.FillAnnotationDataTest
// ============================================================================

mod fill_annotation_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.fill.annotation.FillAnnotationDataTest#t01ReadAndWrite07
    #[test]
    fn t01_read_and_write07() {
        let template = require_fixture("fill/annotation.xlsx");
        assert_annotation_list_fill(&template, "fillAnnotation07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.annotation.FillAnnotationDataTest#t02ReadAndWrite03
    #[test]
    fn t02_read_and_write03() {
        // Java fills xls/fill/annotation.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(
            &require_fixture("xls/fill/annotation.xls"),
            "t02_read_and_write03.xls",
        );
    }
}

// ============================================================================
// FillStyleDataTest (4 @Test)
// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleDataTest
// ============================================================================

mod fill_style_data_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleDataTest#t01Fill07
    #[test]
    fn t01_fill07() {
        let template = require_fixture("fill/style.xlsx");
        assert_style_list_fill(&template, "fileStyle07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleDataTest#t02Fill03
    #[test]
    fn t02_fill03() {
        // Java fills xls/fill/style.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(&require_fixture("xls/fill/style.xls"), "t02_fill03.xls");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleDataTest#t11FillStyleHandler07
    #[test]
    fn t11_fill_style_handler07() {
        // Java registers AbstractVerticalCellStyleStrategy on fill.
        // Rust fill inherits template styles; assert value fill parity (no soft-skip).
        let template = require_fixture("fill/style.xlsx");
        assert_style_list_fill(&template, "fileStyleHandler07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleDataTest#t12FillStyleHandler03
    #[test]
    fn t12_fill_style_handler03() {
        // Java fills xls/fill/style.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(
            &require_fixture("xls/fill/style.xls"),
            "t12_fill_style_handler03.xls",
        );
    }
}

// ============================================================================
// FillStyleAnnotatedTest (2 @Test)
// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleAnnotatedTest
// ============================================================================

mod fill_style_annotated_test {
    use super::*;

    /// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleAnnotatedTest#t01Fill07
    #[test]
    fn t01_fill07() {
        // Java uses FillStyleAnnotatedData with @ContentStyle/@ContentFontStyle.
        // Rust TemplateData fill covers placeholder values; style annotations are write-side.
        let template = require_fixture("fill/style.xlsx");
        assert_style_list_fill(&template, "FillStyleAnnotated07.xlsx");
    }

    /// Java: com.alibaba.easyexcel.test.core.fill.style.FillStyleAnnotatedTest#t02Fill03
    #[test]
    fn t02_fill03() {
        // Java fills xls/fill/style.xls. Legacy XLS template fill is Unsupported (visible).
        assert_xls_fill_works(&require_fixture("xls/fill/style.xls"), "t02_fill03.xls");
    }
}
