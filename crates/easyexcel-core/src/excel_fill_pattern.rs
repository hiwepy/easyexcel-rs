//! Mirrors Java `com.alibaba.excel.enums.poi.FillPatternTypeEnum`.

/// Fill pattern used by annotation-driven cell styles.
///
/// Java retains POI `FillPatternType` codes; Rust strips them. Variant names
/// mirror the POI enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelFillPattern {
    /// No fill.
    None,
    /// Solid foreground fill.
    Solid,
    /// 50% gray pattern.
    MediumGray,
    /// 75% gray pattern.
    DarkGray,
    /// 25% gray pattern.
    LightGray,
    /// Dark horizontal stripes.
    DarkHorizontal,
    /// Dark vertical stripes.
    DarkVertical,
    /// Dark downward diagonal stripes.
    DarkDown,
    /// Dark upward diagonal stripes.
    DarkUp,
    /// Dark grid.
    DarkGrid,
    /// Dark trellis.
    DarkTrellis,
    /// Light horizontal stripes.
    LightHorizontal,
    /// Light vertical stripes.
    LightVertical,
    /// Light downward diagonal stripes.
    LightDown,
    /// Light upward diagonal stripes.
    LightUp,
    /// Light grid.
    LightGrid,
    /// Light trellis.
    LightTrellis,
    /// 12.5% gray pattern.
    Gray125,
    /// 6.25% gray pattern.
    Gray0625,
}
