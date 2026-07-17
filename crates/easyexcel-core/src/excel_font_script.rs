//! Mirrors Java `com.alibaba.excel.enums.poi.FontScript`.

/// Font script position used by annotation-driven font styles.
///
/// Java uses POI `FontScript` codes; Rust strips them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelFontScript {
    /// Normal baseline text.
    None,
    /// Superscript text.
    Superscript,
    /// Subscript text.
    Subscript,
}
