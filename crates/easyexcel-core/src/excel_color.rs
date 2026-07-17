//! Mirrors Java `com.alibaba.excel.enums.poi.IndexedColors` combined with the
//! `0xRRGGBB` extended RGB value.

/// A color supplied by Java-compatible palette index or by explicit RGB value.
///
/// Java uses POI `IndexedColors` short codes; Rust encodes both forms into a
/// single `ExcelColor` enum and offers `java_or_rgb` to apply the historical
/// rule "`<= 64` is a palette index".
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelColor {
    /// A Java EasyExcel / Apache POI indexed palette color. (Java `IndexedColors.shortValue`)
    Indexed(u8),
    /// A backend-neutral RGB color in `0xRRGGBB` form.
    Rgb(u32),
}

impl ExcelColor {
    /// Interprets Java palette indexes `0..=64` as indexed colors and larger values as RGB.
    #[must_use]
    pub const fn java_or_rgb(value: u32) -> Self {
        if value <= 64 {
            // Java palette indexes fit into a single `u8`.
            Self::Indexed(value.to_le_bytes()[0])
        } else {
            Self::Rgb(value)
        }
    }
}
