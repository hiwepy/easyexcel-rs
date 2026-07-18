//! Mirrors Java com.alibaba.excel.util.FileUtils.

#![allow(dead_code)]

use std::fs;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

use tempfile::{NamedTempFile, TempDir};

use crate::excel_error::ExcelError;

static TEMP_FILE_PREFIX: OnceLock<String> = OnceLock::new();
static POI_FILES_PATH: OnceLock<PathBuf> = OnceLock::new();
static CACHE_PATH: OnceLock<PathBuf> = OnceLock::new();

/// Mirrors `org.apache.commons.io.FileUtils#openInputStream`.
pub fn open_input_stream(path: &Path) -> io::Result<std::fs::File> {
    fs::File::open(path)
}

/// Mirrors `com.alibaba.excel.util.FileUtils#writeToFile`.
pub fn write_to_file(path: &Path, data: &[u8]) -> Result<(), ExcelError> {
    let mut file = fs::File::create(path)?;
    file.write_all(data)?;
    Ok(())
}

/// Mirrors `com.alibaba.excel.util.FileUtils#createCacheTmpFile`.
///
/// Uses `tempfile::NamedTempFile` instead of the Java POI cache.
pub fn create_cache_tmp_file() -> io::Result<NamedTempFile> {
    NamedTempFile::new()
}

/// Mirrors `com.alibaba.excel.util.FileUtils#createPoiFilesDirectory`.
///
/// Java creates a `posobody` / `poifiles` temp dir under `java.io.tmpdir`.
/// Rust uses `tempfile::TempDir`.
pub fn create_poi_files_directory() -> io::Result<TempDir> {
    TempDir::new()
}

/// Mirrors `org.apache.commons.io.FileUtils#forceMkdir` / `createDirectory`.
pub fn create_directory(path: &Path) -> Result<(), ExcelError> {
    fs::create_dir_all(path)?;
    Ok(())
}

/// Mirrors `org.apache.commons.io.FileUtils#deleteQuietly` / `delete`.
pub fn delete(path: &Path) -> Result<(), ExcelError> {
    if path.is_dir() {
        fs::remove_dir_all(path)?;
    } else if path.exists() {
        fs::remove_file(path)?;
    }
    Ok(())
}

/// Mirrors `com.alibaba.excel.util.FileUtils#getTempFilePrefix`.
#[must_use]
pub fn get_temp_file_prefix() -> &'static str {
    TEMP_FILE_PREFIX.get_or_init(|| "easyexcel-".to_owned())
}

/// Mirrors `com.alibaba.excel.util.FileUtils#setTempFilePrefix`.
pub fn set_temp_file_prefix(prefix: impl Into<String>) {
    let _ = TEMP_FILE_PREFIX.set(prefix.into());
}

/// Mirrors `com.alibaba.excel.util.FileUtils#getPoiFilesPath`.
#[must_use]
pub fn get_poi_files_path() -> PathBuf {
    POI_FILES_PATH
        .get_or_init(|| std::env::temp_dir().join("easyexcel-poifiles"))
        .clone()
}

/// Mirrors `com.alibaba.excel.util.FileUtils#setPoiFilesPath`.
pub fn set_poi_files_path(path: impl Into<PathBuf>) {
    let _ = POI_FILES_PATH.set(path.into());
}

/// Mirrors `com.alibaba.excel.util.FileUtils#getCachePath`.
#[must_use]
pub fn get_cache_path() -> PathBuf {
    CACHE_PATH
        .get_or_init(|| std::env::temp_dir().join("easyexcel-cache"))
        .clone()
}

/// Mirrors `com.alibaba.excel.util.FileUtils#setCachePath`.
pub fn set_cache_path(path: impl Into<PathBuf>) {
    let _ = CACHE_PATH.set(path.into());
}
