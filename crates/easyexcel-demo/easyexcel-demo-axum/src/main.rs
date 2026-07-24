//! Axum 版 Web 读写演示。
//!
//! 对应 Java `com.alibaba.easyexcel.test.demo.web.WebTest`：
//! - `GET /download`
//! - `GET /downloadFailedUsingJson`
//! - `POST /upload`

use std::net::SocketAddr;

use axum::{
    Router,
    extract::Multipart,
    routing::{get, post},
};
use chrono::NaiveDateTime;
use easyexcel::{AnalysisContext, EasyExcel, ExcelRow, ReadListener, Result as ExcelResult};
use easyexcel_web_axum::{
    ExcelDownloadErrorBody, excel_download_error_response, excel_download_or_json_response,
    excel_download_response, extension_from_path, read_upload_with_listener,
};
use tracing::info;

/// 下载数据行，对应 Java `DownloadData`。
#[derive(Debug, Clone, ExcelRow)]
struct DownloadData {
    #[excel(name = "字符串标题", index = 0)]
    string: String,
    #[excel(name = "日期标题", index = 1)]
    date: NaiveDateTime,
    #[excel(name = "数字标题", index = 2)]
    double_data: f64,
}

/// 上传数据行，对应 Java `UploadData`（无 `@ExcelProperty`，按列序映射）。
#[derive(Debug, Clone, ExcelRow)]
struct UploadData {
    #[excel(index = 0)]
    string: String,
    #[excel(index = 1)]
    date: NaiveDateTime,
    #[excel(index = 2)]
    double_data: f64,
}

/// 批量缓存上传行，对应 Java `UploadDataListener`。
struct UploadDataListener {
    batch: Vec<UploadData>,
}

impl UploadDataListener {
    /// 创建监听器。
    fn new() -> Self {
        Self { batch: Vec::new() }
    }

    /// 模拟 DAO 持久化。
    fn save_batch(rows: &[UploadData]) {
        info!("存储 {} 条上传数据", rows.len());
        for row in rows {
            info!(?row, "解析到一条数据");
        }
    }
}

impl ReadListener<UploadData> for UploadDataListener {
    /// 每解析一行调用一次，对应 Java `invoke`。
    fn invoke(&mut self, data: UploadData, _context: &AnalysisContext) -> ExcelResult<()> {
        self.batch.push(data);
        if self.batch.len() >= 5 {
            Self::save_batch(&self.batch);
            self.batch.clear();
        }
        Ok(())
    }

    /// 全部解析完成后刷盘，对应 Java `doAfterAllAnalysed`。
    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> ExcelResult<()> {
        if !self.batch.is_empty() {
            Self::save_batch(&self.batch);
            self.batch.clear();
        }
        info!("所有数据解析完成！");
        Ok(())
    }
}

/// 构造与 Java WebTest 相同的 10 行样例数据。
fn sample_download_rows() -> Vec<DownloadData> {
    let date = NaiveDateTime::parse_from_str("2020-01-01 00:00:00", "%Y-%m-%d %H:%M:%S")
        .expect("valid demo date");
    (0..10)
        .map(|_| DownloadData {
            string: "字符串0".to_owned(),
            date,
            double_data: 0.56,
        })
        .collect()
}

/// Java 侧 URLEncoder.encode("测试") 的文件名（不含扩展名）。
fn encoded_file_name() -> String {
    urlencoding::encode("测试").replace('+', "%20")
}

/// `GET /download` — 直接返回 XLSX 附件。
async fn download() -> axum::response::Response {
    excel_download_response(&encoded_file_name(), "模板", sample_download_rows()).unwrap_or_else(
        |error| excel_download_error_response(ExcelDownloadErrorBody::download_failed(&error)),
    )
}

/// `GET /downloadFailedUsingJson` — 失败时返回 JSON。
async fn download_failed_using_json() -> axum::response::Response {
    excel_download_or_json_response(&encoded_file_name(), "模板", sample_download_rows())
}

/// `POST /upload` — 读取 multipart 文件并事件解析。
async fn upload(mut multipart: Multipart) -> Result<String, String> {
    while let Some(field) = multipart
        .next_field()
        .await
        .map_err(|error| error.to_string())?
    {
        let file_name = field.file_name().unwrap_or("upload.xlsx").to_owned();
        let bytes = field
            .bytes()
            .await
            .map_err(|error| error.to_string())?
            .to_vec();
        let extension = extension_from_path(std::path::Path::new(&file_name));
        read_upload_with_listener::<UploadData, _>(&bytes, extension, UploadDataListener::new())
            .map_err(|error| error.to_string())?;
        return Ok("success".to_owned());
    }
    Err("未收到上传文件".to_owned())
}

/// 启动 Axum 服务。
#[tokio::main]
async fn main() {
    tracing_subscriber::fmt::init();

    let app = Router::new()
        .route("/download", get(download))
        .route("/downloadFailedUsingJson", get(download_failed_using_json))
        .route("/upload", post(upload));

    let address = SocketAddr::from(([127, 0, 0, 1], 8080));
    info!("Axum WebTest 演示监听 http://{address}");
    info!("GET  /download");
    info!("GET  /downloadFailedUsingJson");
    info!("POST /upload");

    let listener = tokio::net::TcpListener::bind(address).await.expect("bind");
    axum::serve(listener, app).await.expect("serve");
}
