//! Mirrors Java `com.alibaba.excel.annotation.write.style.HeadRowHeight`.
//!
//! In Rust, prefer `#[excel(head_row_height = N)]` on a type with
//! `#[derive(ExcelRow)]`. This marker exists for 1:1 Java package parity.
//!
//! ## Java attributes
//!
//! | Attribute | Default | Meaning |
//! |---|---|---|
//! | `value` | `-1` | Header row height; `-1` means auto height |
//!
//! ## Rust mapping
//!
//! Type-level only → [`crate::ExcelWriteMetadata::head_row_height`].

/// Marker type mirroring Java `@HeadRowHeight`.
///
/// Sets the height of each generated header row. Target is type-level only,
/// matching Java `@Target({ElementType.TYPE})`.
#[allow(dead_code)]
pub struct HeadRowHeight;
