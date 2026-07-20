//! Mirrors Java `com.alibaba.excel.annotation.write.style.ContentStyle`.
//!
//! In Rust, prefer `#[excel(content_style(...))]` on a type or field with
//! `#[derive(ExcelRow)]`. This marker exists for 1:1 Java package parity.
//!
//! ## Java attributes (subset mirrored by [`crate::ExcelCellStyle`])
//!
//! | Attribute | Default | Meaning |
//! |---|---|---|
//! | `dataFormat` | `-1` | Built-in or custom data format |
//! | `hidden` / `locked` / `quotePrefix` | `DEFAULT` | Cell protection / quote-prefix flags |
//! | `horizontalAlignment` / `verticalAlignment` | `DEFAULT` | Alignment enums |
//! | `wrapped` | `DEFAULT` | Text wrap |
//! | `rotation` / `indent` | `-1` | Rotation degrees / indent spaces |
//! | `borderLeft` / `borderRight` / `borderTop` / `borderBottom` | `DEFAULT` | Border styles |
//! | `leftBorderColor` / `rightBorderColor` / `topBorderColor` / `bottomBorderColor` | `-1` | Indexed border colors |
//! | `fillPatternType` | `DEFAULT` | Fill pattern |
//! | `fillBackgroundColor` / `fillForegroundColor` | `-1` | Indexed fill colors |
//! | `shrinkToFit` | `DEFAULT` | Shrink-to-fit flag |
//!
//! ## Rust mapping
//!
//! Nested `#[excel(content_style(wrapped = true, fill_pattern = "solid", ...))]`
//! emits [`crate::ExcelCellStyle`] onto
//! [`crate::ExcelWriteMetadata::content_style`] (type) or
//! [`crate::ExcelColumn::content_style`] (field).

/// Marker type mirroring Java `@ContentStyle`.
///
/// Custom content (data-row) cell styles. Field-level values replace type-level
/// styles for that column.
#[allow(dead_code)]
pub struct ContentStyle;
