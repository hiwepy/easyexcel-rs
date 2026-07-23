//! Axum（Spring Boot）Web 集成层。
//!
//! 对应 Java `com.alibaba.easyexcel.test.demo.web.WebTest` 中的
//! `HttpServletResponse` 下载 / 上传模式。

mod error_body;
mod headers;
mod read_upload;
mod write_response;

#[cfg(test)]
mod tests;

pub use error_body::*;
pub use headers::*;
pub use read_upload::*;
pub use write_response::*;
