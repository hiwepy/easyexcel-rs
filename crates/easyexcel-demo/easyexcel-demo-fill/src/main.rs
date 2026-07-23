//! 模板填充演示。
//!
//! 先写入带占位符的模板，再调用 [`EasyExcel::fill_template`] 生成结果文件。
//!
//! 用法：
//! ```text
//! cargo run -p easyexcel-demo-fill
//! ```

use std::path::PathBuf;

use easyexcel::{EasyExcel, ExcelRow, TemplateData};

/// 用于生成占位符模板的行（写入后手动在 Excel 中通常使用 `{name}` 等占位符；
/// 此处直接写入占位符文本作为单元格值）。
#[derive(Debug, Clone, ExcelRow)]
struct TemplateSeedRow {
    #[excel(name = "姓名", index = 0)]
    label: String,
    #[excel(name = "分数", index = 1)]
    score_label: String,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let template = PathBuf::from("target/demo-fill-template.xlsx");
    let output = PathBuf::from("target/demo-fill-output.xlsx");

    if let Some(parent) = template.parent() {
        std::fs::create_dir_all(parent)?;
    }

    // 写入含占位符的模板（对应 Java fill 示例中的 simple.xlsx 思路）
    EasyExcel::write::<TemplateSeedRow>(&template)
        .sheet("Sheet1")
        .do_write([TemplateSeedRow {
            label: "{name}".to_owned(),
            score_label: "{score}".to_owned(),
        }])?;

    let data = TemplateData::new()
        .with("name", "张三")
        .with("score", 98.5);

    EasyExcel::fill_template(&template, &output, &data)?;

    println!("模板: {}", template.display());
    println!("输出: {}", output.display());
    Ok(())
}
