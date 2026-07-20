//! Temp fill-package *contract* tests — portable EasyExcel fill API.
//!
//! Mirrors Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest`,
//! `temp.FillTempTest`, and `temp.issue1663.FillTest` using fixtures under
//! `tests/fixtures/` (no machine-local paths).

use easyexcel::{
    DynamicRow, DynamicValue, EasyExcel, FillConfig, FillDirection, FillWrapper, TemplateData,
};
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

/// Java `temp.fill.FillTempTest.simpleFill` — map/object fill via TemplateData.
#[test]
fn temp_fill_simple_map_values() {
    let template = fixture("demo/fill/simple.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_simple_map.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三") || dynamic_contains(&rows, "5"),
        "filled values must appear: {rows:?}"
    );
}

/// Java `temp.fill.FillTempTest.listFill` — list template expansion.
#[test]
fn temp_fill_list_template() {
    let template = fixture("demo/fill/list.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_list.xlsx");
    let items: Vec<TemplateData> = (0..10)
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
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.fill.FillTempTest.complexFill` — forceNewRow list fill.
#[test]
fn temp_fill_complex_force_new_row() {
    let template = fixture("demo/fill/complex.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_complex.xlsx");
    let items: Vec<TemplateData> = (0..5)
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
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(
        dynamic_contains(&rows, "张三"),
        "complex fill should contain 张三: {rows:?}"
    );
}

/// Java `temp.fill.FillTempTest.complexFillWithTable`.
#[test]
fn temp_fill_complex_with_table() {
    let template = fixture("demo/fill/complexFillWithTable.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_complex_table.xlsx");
    let items: Vec<TemplateData> = (0..5)
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

/// Java `temp.fill.FillTempTest.horizontalFill`.
#[test]
fn temp_fill_horizontal() {
    let template = fixture("demo/fill/horizontal.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_horizontal.xlsx");
    let items = vec![
        TemplateData::new().with("name", "张三").with("number", 5.2),
        TemplateData::new().with("name", "李四").with("number", 6.2),
    ];
    EasyExcel::fill_template_list(
        &template,
        &output,
        &FillWrapper::new(items),
        FillConfig::new().direction(FillDirection::Horizontal),
    )
    .unwrap();
    assert!(output.exists());
    let bytes = std::fs::read(&output).unwrap();
    assert!(bytes.starts_with(b"PK"));
}

/// Java `temp.fill.FillTempTest.compositeFill`.
#[test]
fn temp_fill_composite() {
    let template = fixture("demo/fill/composite.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_composite.xlsx");
    let data = TemplateData::new()
        .with("name", "张三")
        .with("number", 5.2)
        .with("date", "2019年10月9日");
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}

/// Java `temp.issue1663.FillTest` — named FillWrapper + unknown map key ignored.
#[test]
fn temp_fill_issue1663_named_wrapper() {
    let template = fixture("java/temp/issue1663/template.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_issue1663.xlsx");
    let items: Vec<TemplateData> = (0..10)
        .map(|_| TemplateData::new().with("name", "张三").with("number", 5.2))
        .collect();
    let mut writer = EasyExcel::template_writer(&template, &output).unwrap();
    writer
        .fill_list(
            &FillWrapper::named("data1", items),
            FillConfig::new().direction(FillDirection::Vertical),
        )
        .unwrap();
    // Variable {date} may be absent in template — must not fail.
    writer
        .fill(
            &TemplateData::new()
                .with("date", "2019年10月9日13:28:28")
                .with("total", 1000),
        )
        .unwrap();
    writer.finish().unwrap();
    assert!(output.exists());
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `temp.FillTempTest` intent — style fill template still portable.
#[test]
fn temp_fill_style_template() {
    let template = fixture("fill/style.xlsx");
    assert_fixture(&template);
    let output = temp_path("temp_fill_style.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}
