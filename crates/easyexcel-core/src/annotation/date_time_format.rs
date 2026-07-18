//! Mirrors Java `com.alibaba.excel.annotation.format.DateTimeFormat`.
//!
//! In Rust, `#[excel(format = "...")]` replaces this annotation.
/// Marker type mirroring Java `@DateTimeFormat`.
#[allow(dead_code)]
pub struct DateTimeFormat;
