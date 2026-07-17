//! Mirrors Java `com.alibaba.excel.enums.poi.BorderStyleEnum`.

/// Border line style used by annotation-driven cell styles.
///
/// Java retains the POI `BorderStyle` codes; Rust drops them because the
/// underlying backend does not need them. Variant names match the Java
/// POI enum names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelBorderStyle {
    /// No border.
    None,
    /// Thin solid line.
    Thin,
    /// Medium solid line.
    Medium,
    /// Dashed line.
    Dashed,
    /// Dotted line.
    Dotted,
    /// Thick solid line.
    Thick,
    /// Double line.
    Double,
    /// Hairline border.
    Hair,
    /// Medium dashed line.
    MediumDashed,
    /// Dash-dot line.
    DashDot,
    /// Medium dash-dot line.
    MediumDashDot,
    /// Dash-dot-dot line.
    DashDotDot,
    /// Medium dash-dot-dot line.
    MediumDashDotDot,
    /// Slanted dash-dot line.
    SlantDashDot,
}
