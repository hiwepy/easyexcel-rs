//! Web 下载失败时的 JSON 响应体。
//!
//! 对应 Java `WebTest.downloadFailedUsingJson` 中 Fastjson 序列化的
//! `Map<String, String>`（`status` / `message` 键）。

use serde::{Deserialize, Serialize};

/// Excel 下载失败时返回的 JSON 体。
///
/// 对应 Java：
/// ```java
/// map.put("status", "failure");
/// map.put("message", "下载文件失败" + e.getMessage());
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ExcelDownloadErrorBody {
    /// 固定为 `"failure"`，与 Java WebTest 一致。
    pub status: String,
    /// 人类可读的错误说明，通常以「下载文件失败」为前缀。
    pub message: String,
}

impl ExcelDownloadErrorBody {
    /// 构造 Java WebTest 兼容的失败响应体。
    ///
    /// 对应 Java `map.put("status", "failure")` +
    /// `map.put("message", "下载文件失败" + e.getMessage())`。
    #[must_use]
    pub fn failure(message: impl Into<String>) -> Self {
        Self {
            status: "failure".to_owned(),
            message: message.into(),
        }
    }

    /// 使用 Java WebTest 默认前缀包装底层错误信息。
    #[must_use]
    pub fn download_failed(error: impl std::fmt::Display) -> Self {
        Self::failure(format!("下载文件失败{error}"))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// 与 Java WebTest Fastjson 输出键名对齐的契约测试。
    #[test]
    fn json_keys_match_java_web_test() {
        let body = ExcelDownloadErrorBody::download_failed("boom");
        let json = serde_json::to_string(&body).expect("serialize");
        assert!(json.contains("\"status\":\"failure\""));
        assert!(json.contains("\"message\":\"下载文件失败boom\""));
        let back: ExcelDownloadErrorBody = serde_json::from_str(&json).expect("deserialize");
        assert_eq!(back, body);
    }
}
