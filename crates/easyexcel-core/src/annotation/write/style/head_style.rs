//! Mirrors Java `com.alibaba.excel.annotation.write.style.HeadStyle`.
//!
//! In Rust, prefer `#[excel(head_style(...))]` on a type or field with
//! `#[derive(ExcelRow)]`. This marker exists for 1:1 Java package parity.
//!
//! ## Java attributes (subset mirrored by [`crate::ExcelCellStyle`])
//!
//! Same cell-style inventory as [`super::content_style::ContentStyle`]:
//! `dataFormat`, `hidden`, `locked`, `quotePrefix`, alignments, `wrapped`,
//! `rotation`, `indent`, borders/colors, fill pattern/colors, and `shrinkToFit`.
//!
//! ## Rust mapping
//!
//! Nested `#[excel(head_style(horizontal_alignment = "center", ...))]`
//! emits [`crate::ExcelCellStyle`] onto
//! [`crate::ExcelWriteMetadata::head_style`] (type) or
//! [`crate::ExcelColumn::head_style`] (field).

/// Marker type mirroring Java `@HeadStyle`.
///
/// Custom header cell styles. Field-level values replace type-level styles for
/// that column's header cell.
#[allow(dead_code)]
pub struct HeadStyle;
