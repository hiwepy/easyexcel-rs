//! Mirrors Java `com.alibaba.excel.annotation.write.style.OnceAbsoluteMerge`.
//!
//! In Rust, prefer
//! `#[excel(once_absolute_merge(first_row_index = ..., last_row_index = ..., first_column_index = ..., last_column_index = ...))]`
//! on a type with `#[derive(ExcelRow)]`. This marker exists for 1:1 Java
//! package parity.
//!
//! ## Java attributes
//!
//! | Attribute | Default | Meaning |
//! |---|---|---|
//! | `firstRowIndex` | `-1` | Inclusive first row (zero-based) |
//! | `lastRowIndex` | `-1` | Inclusive last row (zero-based) |
//! | `firstColumnIndex` | `-1` | Inclusive first column (zero-based) |
//! | `lastColumnIndex` | `-1` | Inclusive last column (zero-based) |
//!
//! Java builds [`crate::OnceAbsoluteMergeProperty`] and registers an
//! `OnceAbsoluteMergeStrategy` once per sheet.
//!
//! ## Rust mapping
//!
//! Type-level only → [`crate::ExcelWriteMetadata::once_absolute_merge`] as
//! [`crate::OnceAbsoluteMergeProperty`]. Negative defaults are treated as
//! unset and skipped by the writer.

/// Marker type mirroring Java `@OnceAbsoluteMerge`.
///
/// Merges an absolute cell region once when the sheet is created. Target is
/// type-level only, matching Java `@Target({ElementType.TYPE})`.
#[allow(dead_code)]
pub struct OnceAbsoluteMerge;
