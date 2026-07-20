//! Shared helpers for temp package 1:1 method-name matrix tests.

mod fill_assert;
mod io;

pub use fill_assert::*;
pub use io::*;

use easyexcel::DynamicRow;
use easyexcel::DynamicValue;
use tempfile::tempdir;

/// Build a unique temp file path under a fresh tempdir.
pub fn temp_path(name: &str) -> std::path::PathBuf {
    tempdir().unwrap().keep().join(name)
}

/// Resolve a path under `tests/fixtures/`.
pub fn fixture(name: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name)
}

/// Assert a Java-ported fixture exists.
pub fn assert_fixture(path: &std::path::Path) {
    assert!(
        path.exists(),
        "required Java fixture missing: {}",
        path.display()
    );
}

/// Whether any dynamic cell string contains `needle`.
pub fn dynamic_contains(rows: &[DynamicRow], needle: &str) -> bool {
    rows.iter().any(|row| {
        row.values().iter().any(|(_, val)| match val {
            DynamicValue::String(s) => s.contains(needle),
            DynamicValue::ActualData(easyexcel::CellValue::String(s)) => s.contains(needle),
            _ => false,
        })
    })
}
