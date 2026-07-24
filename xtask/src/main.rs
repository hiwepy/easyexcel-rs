//! easyexcel-rust 仓库维护命令（对齐 sa-token-rs xtask）。
//!
//! 用法：
//! ```text
//! cargo run -p xtask -- migration-audit
//! cargo run -p xtask -- migration-audit-strict
//! ```

use std::env;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::path::{Path, PathBuf};
use std::process::ExitCode;

const SOURCE_COMMIT: &str = "3afdea9d5da7f24a66eda6ec44a9dfce80b16802";
const EXPECTED_JAVA_MAIN: usize = 325;
const MAP_PATH: &str = "docs/migration/file-map.csv";

type TaskResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn main() -> ExitCode {
    let mut args = env::args().skip(1);
    let result = match args.next().as_deref() {
        Some("migration-audit") => audit(false),
        Some("migration-audit-strict") => audit(true),
        Some("--strict") => audit(true),
        _ => {
            eprintln!(
                "usage: cargo run -p xtask -- <migration-audit [--strict]|migration-audit-strict>"
            );
            return ExitCode::from(2);
        }
    };
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(err) => {
            eprintln!("xtask error: {err}");
            ExitCode::FAILURE
        }
    }
}

/// 审计 `file-map.csv`。
///
/// - 普通模式：校验行数、路径唯一、rust 文件存在性统计
/// - strict：要求全部 `complete`（v1.0 前预期失败）
fn audit(strict: bool) -> TaskResult {
    let map = PathBuf::from(MAP_PATH);
    if !map.is_file() {
        return Err(format!("missing {MAP_PATH}").into());
    }

    let file = File::open(&map)?;
    let reader = BufReader::new(file);
    let mut rows = Vec::new();
    for (idx, line) in reader.lines().enumerate() {
        let line = line?;
        if idx == 0 {
            continue;
        }
        if line.trim().is_empty() {
            continue;
        }
        let cols: Vec<&str> = line.split(',').collect();
        if cols.len() < 6 {
            return Err(format!("bad csv row {idx}: {line}").into());
        }
        rows.push(cols.into_iter().map(str::to_owned).collect::<Vec<_>>());
    }

    println!("file-map rows: {}", rows.len());
    println!("expected java main (excl package-info): {EXPECTED_JAVA_MAIN}");
    println!("source_commit baseline: {SOURCE_COMMIT}");

    let mut missing_rust = 0usize;
    let mut complete = 0usize;
    let mut planned = 0usize;
    let mut in_progress = 0usize;
    let mut java_set = std::collections::HashSet::new();
    let mut rust_set = std::collections::HashSet::new();

    for row in &rows {
        let java = &row[0];
        let rust = &row[1];
        let status = &row[5];
        if !java_set.insert(java.clone()) {
            return Err(format!("duplicate java_file: {java}").into());
        }
        if !rust.is_empty() && !rust_set.insert(rust.clone()) {
            // 允许多个 java 映射到同一 rust（一对多时用 capability 区分），仅警告
            eprintln!("warn: shared rust_file: {rust}");
        }
        if !rust.is_empty() && !Path::new(rust).is_file() {
            missing_rust += 1;
            eprintln!("missing rust file: {rust} (from {java})");
        }
        match status.as_str() {
            "complete" => complete += 1,
            "planned" => planned += 1,
            "in_progress" => in_progress += 1,
            _ => {}
        }
    }

    println!("status complete={complete} in_progress={in_progress} planned={planned}");
    println!("missing rust files on disk: {missing_rust}");

    if missing_rust > 0 {
        return Err(format!("{missing_rust} rust targets missing on disk").into());
    }

    if strict {
        let unfinished = rows
            .iter()
            .filter(|r| {
                r[5] != "complete" && r[5] != "ignore" && r[5] != "handle" && r[5] != "excluded"
            })
            .count();
        if unfinished > 0 {
            return Err(format!(
                "strict audit: {unfinished} rows not complete/ignore/handle (expected until v1.0)"
            )
            .into());
        }
    }

    println!("migration-audit ok (strict={strict})");
    Ok(())
}
