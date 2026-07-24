//! 上传 Excel 读取辅助（Actix-web 版）。

use std::io::Write;
use std::path::{Path, PathBuf};

use easyexcel::{EasyExcel, ExcelRow, ReadListener};
use easyexcel_core::Result;
use tempfile::Builder;

/// 将上传字节写入临时文件。
///
/// # Errors
///
/// 临时文件创建或写入失败时返回 I/O 错误。
pub fn write_upload_temp(
    bytes: &[u8],
    extension: &str,
) -> Result<(PathBuf, tempfile::NamedTempFile)> {
    let suffix = normalize_extension(extension);
    let mut temp = Builder::new().suffix(&suffix).tempfile()?;
    temp.as_file_mut().write_all(bytes)?;
    temp.as_file().sync_all()?;
    Ok((temp.path().to_path_buf(), temp))
}

/// 事件驱动读取上传内容。
///
/// # Errors
///
/// 解析或监听器错误时返回。
pub fn read_upload_with_listener<T, L>(bytes: &[u8], extension: &str, listener: L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let (path, _temp) = write_upload_temp(bytes, extension)?;
    EasyExcel::read::<T, L>(path, listener).do_read()
}

/// 同步收集上传 Excel 的全部行。
///
/// # Errors
///
/// 解析或行转换失败时返回。
pub fn read_upload_sync<T>(bytes: &[u8], extension: &str) -> Result<Vec<T>>
where
    T: ExcelRow,
{
    let (path, _temp) = write_upload_temp(bytes, extension)?;
    EasyExcel::read_sync::<T>(&path).do_read_sync()
}

/// 从路径推断扩展名；缺省为 `.xlsx`。
#[must_use]
pub fn extension_from_path(path: &Path) -> &'static str {
    path.extension()
        .and_then(|value| value.to_str())
        .map(|value| match value.to_ascii_lowercase().as_str() {
            "csv" => ".csv",
            "xls" => ".xls",
            _ => ".xlsx",
        })
        .unwrap_or(".xlsx")
}

fn normalize_extension(extension: &str) -> String {
    if extension.starts_with('.') {
        extension.to_owned()
    } else {
        format!(".{extension}")
    }
}
