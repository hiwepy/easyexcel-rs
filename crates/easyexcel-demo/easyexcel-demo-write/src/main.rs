//! 简单写入演示：生成 XLSX 或 CSV。
//!
//! 用法：
//! ```text
//! cargo run -p easyexcel-demo-write -- target/demo-write.xlsx
//! ```

use std::env;
use std::path::PathBuf;

use chrono::NaiveDateTime;
use easyexcel::{EasyExcel, ExcelRow};

/// 演示用行模型。
#[derive(Debug, Clone, ExcelRow)]
struct DemoRow {
    #[excel(name = "名称", index = 0)]
    name: String,
    #[excel(name = "日期", index = 1)]
    date: NaiveDateTime,
    #[excel(name = "数值", index = 2)]
    amount: f64,
}

fn sample_rows() -> Vec<DemoRow> {
    let date = NaiveDateTime::parse_from_str("2024-06-01 12:00:00", "%Y-%m-%d %H:%M:%S")
        .expect("valid date");
    (0..5)
        .map(|index| DemoRow {
            name: format!("项目{index}"),
            date,
            amount: f64::from(index) + 0.5,
        })
        .collect()
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/demo-write.xlsx"));

    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }

    EasyExcel::write::<DemoRow>(&path)
        .sheet("数据")
        .do_write(sample_rows())?;

    println!("已写入 {} 行到 {}", 5, path.display());
    Ok(())
}
