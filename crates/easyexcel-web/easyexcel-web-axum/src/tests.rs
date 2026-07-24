//! Web JSON 与 Java WebTest 的 serde 键名对齐测试。

use easyexcel_core::ExcelDownloadErrorBody;
use serde_json::{Value, json};

#[test]
fn excel_download_error_body_serializes_java_keys() {
    let body = ExcelDownloadErrorBody::download_failed("stream closed");
    let value: Value = serde_json::to_value(&body).expect("serialize");
    assert_eq!(value["status"], json!("failure"));
    assert_eq!(value["message"], json!("下载文件失败stream closed"));
    assert!(value.get("status").is_some());
    assert!(value.get("message").is_some());
}

#[test]
fn excel_download_error_body_deserializes_java_keys() {
    let raw = r#"{"status":"failure","message":"下载文件失败测试"}"#;
    let body: ExcelDownloadErrorBody = serde_json::from_str(raw).expect("deserialize");
    assert_eq!(body.status, "failure");
    assert_eq!(body.message, "下载文件失败测试");
}
