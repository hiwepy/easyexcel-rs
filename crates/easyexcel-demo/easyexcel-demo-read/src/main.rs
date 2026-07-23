//! 简单读取演示：XLSX 与 CSV。
//!
//! 用法：
//! ```text
//! cargo run -p easyexcel-demo-read -- target/demo-read.xlsx
//! cargo run -p easyexcel-demo-read -- target/demo-read.csv
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

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("target/demo-read.xlsx"));

    let rows = EasyExcel::read_sync::<DemoRow>(&path).do_read_sync()?;
    println!("读取 {} 行自 {}", rows.len(), path.display());
    for (index, row) in rows.iter().enumerate() {
        println!("{index}: {row:?}");
    }
    Ok(())
}
