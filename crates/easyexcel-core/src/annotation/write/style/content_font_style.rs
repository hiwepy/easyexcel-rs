//! Mirrors Java `com.alibaba.excel.annotation.write.style.ContentFontStyle`.
//!
//! In Rust, prefer `#[excel(content_font_style(...))]` on a type or field with
//! `#[derive(ExcelRow)]`. This marker exists for 1:1 Java package parity.
//!
//! ## Java attributes (subset mirrored by [`crate::ExcelFontStyle`])
//!
//! | Attribute | Default | Meaning |
//! |---|---|---|
//! | `fontName` | `""` | Font family (e.g. `"Arial"`) |
//! | `fontHeightInPoints` | `-1` | Font size in points |
//! | `italic` / `strikeout` / `bold` | `DEFAULT` | Font face flags |
//! | `color` | `-1` | Indexed or palette color |
//! | `typeOffset` | `-1` | Super/subscript (`SS_NONE` / `SS_SUPER` / `SS_SUB`) |
//! | `underline` | `-1` | Underline style byte |
//! | `charset` | `-1` | Character set |
//!
//! ## Rust mapping
//!
//! Nested `#[excel(content_font_style(bold = true, italic = true, ...))]`
//! emits [`crate::ExcelFontStyle`] onto
//! [`crate::ExcelWriteMetadata::content_font_style`] (type) or
//! [`crate::ExcelColumn::content_font_style`] (field).

/// Marker type mirroring Java `@ContentFontStyle`.
///
/// Custom content (data-row) font styles. Independent from cell-style
/// inheritance: font and cell styles can be set separately.
#[allow(dead_code)]
pub struct ContentFontStyle;
