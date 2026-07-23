//! Excel 写入与 Actix-web 响应构建。

use actix_web::HttpResponse;
use easyexcel::{EasyExcel, ExcelRow};
use easyexcel_core::{ExcelDownloadErrorBody, Result};
use serde_json;

use crate::headers::excel_xlsx_attachment_headers;

/// 将 [`ExcelRow`] 行序列化为 XLSX 字节数组。
///
/// # Errors
///
/// 行转换或 OOXML 写入失败时返回 [`easyexcel_core::ExcelError`]。
pub fn write_rows_to_bytes<T, I>(sheet_name: &str, rows: I) -> Result<Vec<u8>>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let mut buffer = Vec::new();
    EasyExcel::write::<T>("download.xlsx")
        .sheet(sheet_name)
        .to_writer(&mut buffer)
        .do_write(rows)?;
    Ok(buffer)
}

/// 由 XLSX 字节构建 Actix 附件响应。
#[must_use]
pub fn excel_download_response_from_bytes(file_name: &str, bytes: Vec<u8>) -> HttpResponse {
    let (content_type, content_disposition) =
        crate::headers::excel_xlsx_attachment_headers(file_name);
    HttpResponse::Ok()
        .insert_header((actix_web::http::header::CONTENT_TYPE, content_type))
        .insert_header((actix_web::http::header::CONTENT_DISPOSITION, content_disposition))
        .body(bytes)
}

/// 一步完成写入并返回 Actix XLSX 附件响应。
///
/// # Errors
///
/// 写入失败时返回错误（由调用方决定降级策略）。
pub fn excel_download_response<T, I>(
    file_name: &str,
    sheet_name: &str,
    rows: I,
) -> Result<HttpResponse>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let bytes = write_rows_to_bytes::<T, _>(sheet_name, rows)?;
    Ok(excel_download_response_from_bytes(file_name, bytes))
}

/// 下载失败时返回 JSON 体。
#[must_use]
pub fn excel_download_error_response(body: ExcelDownloadErrorBody) -> HttpResponse {
    HttpResponse::InternalServerError()
        .content_type("application/json; charset=utf-8")
        .body(
            serde_json::to_string(&body).unwrap_or_else(|_| {
                r#"{"status":"failure","message":"下载文件失败JSON序列化错误"}"#.to_owned()
            }),
        )
}

/// 尝试生成 XLSX 附件；失败时自动降级为 JSON 错误体。
#[must_use]
pub fn excel_download_or_json_response<T, I>(
    file_name: &str,
    sheet_name: &str,
    rows: I,
) -> HttpResponse
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    match write_rows_to_bytes::<T, _>(sheet_name, rows) {
        Ok(bytes) => excel_download_response_from_bytes(file_name, bytes),
        Err(error) => excel_download_error_response(ExcelDownloadErrorBody::download_failed(&error)),
    }
}
