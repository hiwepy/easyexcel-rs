//! Mirrors Java `com.alibaba.excel.enums.poi.HorizontalAlignmentEnum`.

/// Horizontal alignment used by annotation-driven cell styles.
///
/// Java retains the POI alignment codes; Rust strips them because the
/// underlying backend (`rust_xlsxwriter`) does not need them. Variant names
/// match Java's `HorizontalAlignment`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelHorizontalAlignment {
    /// Excel's type-dependent default.
    General,
    /// Left aligned.
    Left,
    /// Centered.
    Center,
    /// Right aligned.
    Right,
    /// Repeats content across the cell.
    Fill,
    /// Justified.
    Justify,
    /// Centered across adjacent cells.
    CenterAcross,
    /// Distributed across the cell.
    Distributed,
}
