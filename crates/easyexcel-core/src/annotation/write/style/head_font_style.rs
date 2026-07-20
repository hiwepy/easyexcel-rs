//! Mirrors Java `com.alibaba.excel.annotation.write.style.HeadFontStyle`.
//!
//! In Rust, prefer `#[excel(head_font_style(...))]` on a type or field with
//! `#[derive(ExcelRow)]`. This marker exists for 1:1 Java package parity.
//!
//! ## Java attributes (subset mirrored by [`crate::ExcelFontStyle`])
//!
//! Same font inventory as [`super::content_font_style::ContentFontStyle`]:
//! `fontName`, `fontHeightInPoints`, `italic`, `strikeout`, `color`,
//! `typeOffset`, `underline`, `charset`, and `bold`.
//!
//! ## Rust mapping
//!
//! Nested `#[excel(head_font_style(font_name = "Arial", bold = true, ...))]`
//! emits [`crate::ExcelFontStyle`] onto
//! [`crate::ExcelWriteMetadata::head_font_style`] (type) or
//! [`crate::ExcelColumn::head_font_style`] (field).

/// Marker type mirroring Java `@HeadFontStyle`.
///
/// Custom header font styles. Field-level values replace type-level fonts for
/// that column's header cell.
#[allow(dead_code)]
pub struct HeadFontStyle;
