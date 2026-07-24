//! Actix-web 版 Web 读写演示。

use actix_multipart::Multipart;
use actix_web::{App, HttpResponse, HttpServer, Responder, web};
use chrono::NaiveDateTime;
use easyexcel::{AnalysisContext, ExcelRow, ReadListener, Result as ExcelResult};
use easyexcel_web_actix::{
    excel_download_or_json_response, excel_download_response, extension_from_path,
    read_upload_with_listener,
};
use futures_util::StreamExt;
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

/// 上传数据行，对应 Java `UploadData`。
#[derive(Debug, Clone, ExcelRow)]
struct UploadData {
    #[excel(index = 0)]
    string: String,
    #[excel(index = 1)]
    date: NaiveDateTime,
    #[excel(index = 2)]
    double_data: f64,
}

/// 上传监听器，对应 Java `UploadDataListener`。
struct UploadDataListener {
    batch: Vec<UploadData>,
}

impl UploadDataListener {
    fn new() -> Self {
        Self { batch: Vec::new() }
    }

    fn save_batch(rows: &[UploadData]) {
        info!("存储 {} 条上传数据", rows.len());
        for row in rows {
            info!(?row, "解析到一条数据");
        }
    }
}

impl ReadListener<UploadData> for UploadDataListener {
    fn invoke(&mut self, data: UploadData, _context: &AnalysisContext) -> ExcelResult<()> {
        self.batch.push(data);
        if self.batch.len() >= 5 {
            Self::save_batch(&self.batch);
            self.batch.clear();
        }
        Ok(())
    }

    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> ExcelResult<()> {
        if !self.batch.is_empty() {
            Self::save_batch(&self.batch);
            self.batch.clear();
        }
        info!("所有数据解析完成！");
        Ok(())
    }
}

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

fn encoded_file_name() -> String {
    urlencoding::encode("测试").replace('+', "%20")
}

async fn download() -> impl Responder {
    match excel_download_response(&encoded_file_name(), "模板", sample_download_rows()) {
        Ok(response) => response,
        Err(error) => HttpResponse::InternalServerError().body(error.to_string()),
    }
}

async fn download_failed_using_json() -> impl Responder {
    excel_download_or_json_response(&encoded_file_name(), "模板", sample_download_rows())
}

async fn upload(mut payload: Multipart) -> impl Responder {
    // 对应 Java `WebTest.upload(MultipartFile file)`。
    while let Some(item) = payload.next().await {
        let mut field = match item {
            Ok(field) => field,
            Err(error) => return HttpResponse::BadRequest().body(error.to_string()),
        };
        let file_name = field
            .content_disposition()
            .and_then(|value| value.get_filename().map(str::to_owned))
            .unwrap_or_else(|| "upload.xlsx".to_owned());
        let mut bytes = Vec::new();
        while let Some(chunk) = field.next().await {
            match chunk {
                Ok(data) => bytes.extend_from_slice(data.as_ref()),
                Err(error) => return HttpResponse::BadRequest().body(error.to_string()),
            }
        }
        let extension = extension_from_path(std::path::Path::new(&file_name));
        if let Err(error) =
            read_upload_with_listener::<UploadData, _>(&bytes, extension, UploadDataListener::new())
        {
            return HttpResponse::InternalServerError().body(error.to_string());
        }
        return HttpResponse::Ok().body("success");
    }
    HttpResponse::BadRequest().body("未收到上传文件")
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    tracing_subscriber::fmt::init();
    info!("Actix WebTest 演示监听 http://127.0.0.1:8081");

    HttpServer::new(|| {
        App::new()
            .route("/download", web::get().to(download))
            .route(
                "/downloadFailedUsingJson",
                web::get().to(download_failed_using_json),
            )
            .route("/upload", web::post().to(upload))
    })
    .bind(("127.0.0.1", 8081))?
    .run()
    .await
}
