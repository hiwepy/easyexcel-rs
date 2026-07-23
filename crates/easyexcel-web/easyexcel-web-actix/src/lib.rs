//! Actix-web（Quarkus）Web 集成层。
//!
//! API 与 [`easyexcel-web-axum`] 对称，对应同一套 Java `WebTest` 示例。

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
