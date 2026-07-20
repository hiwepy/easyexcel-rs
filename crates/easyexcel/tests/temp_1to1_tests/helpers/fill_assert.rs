//! Fill-related portable asserts for temp 1:1 matrix.

use easyexcel::{EasyExcel, FillConfig, FillDirection, FillWrapper, TemplateData};

use super::{assert_fixture, dynamic_contains, fixture, temp_path};

pub fn assert_fill_simple() {
    let template = fixture("demo/fill/simple.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_fill_simple.xlsx");
    let data = TemplateData::new().with("name", "张三").with("number", 5.2);
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    let rows = EasyExcel::read_dynamic_sync(&output)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
    assert!(dynamic_contains(&rows, "张三") || dynamic_contains(&rows, "5"));
}

/// List template fill.
pub fn assert_fill_list() {
    let template = fixture("demo/fill/list.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_fill_list.xlsx");
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
}

/// Complex forceNewRow fill.
pub fn assert_fill_complex() {
    let template = fixture("demo/fill/complex.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_fill_complex.xlsx");
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
}

/// Complex fill with table template.
pub fn assert_fill_table() {
    let template = fixture("demo/fill/complexFillWithTable.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_fill_table.xlsx");
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

/// Horizontal fill.
pub fn assert_fill_horizontal() {
    let template = fixture("demo/fill/horizontal.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_fill_horizontal.xlsx");
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
}

/// Composite fill.
pub fn assert_fill_composite() {
    let template = fixture("demo/fill/composite.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_fill_composite.xlsx");
    let data = TemplateData::new()
        .with("name", "张三")
        .with("number", 5.2)
        .with("date", "2019年10月9日");
    EasyExcel::fill_template(&template, &output, &data).unwrap();
    assert!(output.exists());
}

/// issue1663 named FillWrapper.
pub fn assert_fill_issue1663() {
    let template = fixture("java/temp/issue1663/template.xlsx");
    assert_fixture(&template);
    let output = temp_path("1to1_issue1663.xlsx");
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
    writer
        .fill(
            &TemplateData::new()
                .with("date", "2019年10月9日13:28:28")
                .with("total", 1000),
        )
        .unwrap();
    writer.finish().unwrap();
    assert!(output.exists());
}
