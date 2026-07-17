//! Mirrors Java `com.alibaba.excel.enums.HeadKindEnum`.

/// The types of header.
///
/// Rust port of Java `HeadKindEnum`. Distinguishes no-header, class-driven
/// headers, and ad-hoc string-list headers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeadKind {
    /// No header configured.
    None,
    /// Header derived from a `#[derive(ExcelRow)]` class.
    Class,
    /// Header derived from a literal string list.
    String,
}
