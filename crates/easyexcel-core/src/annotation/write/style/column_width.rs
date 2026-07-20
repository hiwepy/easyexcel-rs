//! Mirrors Java `com.alibaba.excel.annotation.write.style.ColumnWidth`.
//!
//! In Rust, prefer `#[excel(column_width = N)]` on a type or field with
//! `#[derive(ExcelRow)]`. This marker exists for 1:1 Java package parity.
//!
//! ## Java attributes
//!
//! | Attribute | Default | Meaning |
//! |---|---|---|
//! | `value` | `-1` | Column width in Excel character units; `-1` keeps the workbook default |
//!
//! ## Rust mapping
//!
//! - Type-level → [`crate::ExcelWriteMetadata::column_width`]
//! - Field-level → [`crate::ExcelColumn::column_width`] (overrides type width)

/// Marker type mirroring Java `@ColumnWidth`.
///
/// Sets the width of a table column. Apply at type scope for a default width,
/// or at field scope to override a single column.
#[allow(dead_code)]
pub struct ColumnWidth;
