//! Mirrors Java `com.alibaba.excel.annotation.write.style.ContentLoopMerge`.
//!
//! In Rust, prefer `#[excel(content_loop_merge(each_row = N, column_extend = M))]`
//! on a field with `#[derive(ExcelRow)]`. This marker exists for 1:1 Java
//! package parity.
//!
//! ## Java attributes
//!
//! | Attribute | Default | Meaning |
//! |---|---|---|
//! | `eachRow` | `1` | Number of data rows in each merge group |
//! | `columnExtend` | `1` | Number of columns spanned by each merge |
//!
//! Java builds [`crate::LoopMergeProperty`] and registers a
//! `LoopMergeStrategy` with the field's physical column index.
//!
//! ## Rust mapping
//!
//! Field-level only → [`crate::ExcelColumn::loop_merge`] as
//! [`crate::LoopMergeProperty`]. The writer converts it into a
//! repeating merge using the column's physical index.

/// Marker type mirroring Java `@ContentLoopMerge`.
///
/// Declares repeating merge regions for content rows of a single field.
/// Target is field-level only, matching Java `@Target({ElementType.FIELD})`.
#[allow(dead_code)]
pub struct ContentLoopMerge;
