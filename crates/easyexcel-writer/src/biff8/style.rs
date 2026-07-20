//! BIFF8 workbook-level FONT / XF / palette registry.
//!
//! Maps Java EasyExcel / POI HSSF style knobs (`WriteCellStyle`, `WriteFont`,
//! `IndexedColors`, `CellStyle` builder) onto FONT + XF records. Borders,
//! custom number formats beyond date/datetime, rich-text runs, and conditional
//! formatting remain unsupported for Minimal BIFF8.

use std::collections::HashMap;

use easyexcel_core::{
    ExcelCellStyle, ExcelColor, ExcelFillPattern, ExcelFontStyle, ExcelHorizontalAlignment,
    ExcelVerticalAlignment,
};

use super::encode::{
    pack_cell_xf, pack_font, ICV_AUTO, ICV_PATTERN_BG_DEFAULT, XF_CUSTOM_BASE, XF_DATE,
    XF_DATETIME, XF_GENERAL,
};

/// Resolved write-style inputs used when allocating an XF index.
#[derive(Debug, Clone, Default)]
pub struct Biff8StyleRequest {
    /// Bold font.
    pub bold: bool,
    /// Italic font.
    pub italic: bool,
    /// Strike-through font.
    pub strikeout: bool,
    /// Font height in points (`None` → 10pt Arial default).
    pub font_height_points: Option<u16>,
    /// Font family name (`None` → `"Arial"`).
    pub font_name: Option<String>,
    /// Font colour as palette ICV (`None` → automatic).
    pub font_color_icv: Option<u16>,
    /// Horizontal alignment POI code (`None` → general / 0).
    pub halign: Option<u8>,
    /// Vertical alignment POI code (`None` → bottom / 2).
    pub valign: Option<u8>,
    /// Wrap text.
    pub wrap: bool,
    /// Fill pattern POI code (`None` / 0 → no fill).
    pub fill_pattern: Option<u8>,
    /// Fill foreground palette ICV.
    pub fill_fg_icv: Option<u16>,
    /// Fill background palette ICV.
    pub fill_bg_icv: Option<u16>,
}

impl Biff8StyleRequest {
    /// Returns `true` when this request would produce XF_GENERAL with default font.
    #[must_use]
    pub fn is_default(&self) -> bool {
        !self.bold
            && !self.italic
            && !self.strikeout
            && self.font_height_points.is_none()
            && self.font_name.is_none()
            && self.font_color_icv.is_none()
            && self.halign.is_none()
            && self.valign.is_none()
            && !self.wrap
            && self.fill_pattern.unwrap_or(0) == 0
            && self.fill_fg_icv.is_none()
    }

    /// Merges annotation / strategy [`ExcelCellStyle`] (Java `WriteCellStyle`).
    pub fn apply_excel_cell_style(&mut self, style: ExcelCellStyle) {
        if let Some(align) = style.horizontal_alignment {
            self.halign = Some(excel_halign(align));
        }
        if let Some(align) = style.vertical_alignment {
            self.valign = Some(excel_valign(align));
        }
        if let Some(wrapped) = style.wrapped {
            self.wrap = wrapped;
        }
        if let Some(pattern) = style.fill_pattern {
            self.fill_pattern = Some(excel_fill_pattern(pattern));
        }
        if let Some(color) = style.fill_foreground_color {
            self.fill_fg_icv = Some(rgb_or_indexed_to_icv(color));
            if self.fill_pattern.unwrap_or(0) == 0 {
                self.fill_pattern = Some(1);
            }
        }
        if let Some(color) = style.fill_background_color {
            self.fill_bg_icv = Some(rgb_or_indexed_to_icv(color));
        }
        if let Some(font) = style.font {
            self.apply_excel_font_style(font);
        }
    }

    /// Merges annotation / strategy [`ExcelFontStyle`] (Java `WriteFont`).
    pub fn apply_excel_font_style(&mut self, style: ExcelFontStyle) {
        if let Some(name) = style.font_name {
            self.font_name = Some(name.to_owned());
        }
        if let Some(height) = style.font_height_in_points {
            #[allow(clippy::cast_possible_truncation, clippy::cast_sign_loss)]
            {
                self.font_height_points = Some(height.round().clamp(1.0, 409.0) as u16);
            }
        }
        if let Some(italic) = style.italic {
            self.italic = italic;
        }
        if let Some(strikeout) = style.strikeout {
            self.strikeout = strikeout;
        }
        if let Some(bold) = style.bold {
            self.bold = bold;
        }
        if let Some(color) = style.color {
            self.font_color_icv = Some(rgb_or_indexed_to_icv(color));
        }
    }
}

#[derive(Debug, Clone, Hash, Eq, PartialEq)]
struct FontKey {
    height_points: u16,
    bold: bool,
    italic: bool,
    strikeout: bool,
    color_icv: u16,
    name: String,
}

#[derive(Debug, Clone, Hash, Eq, PartialEq)]
struct XfKey {
    font_index: u16,
    ifmt: u16,
    halign: u8,
    valign: u8,
    wrap: bool,
    fill_pattern: u8,
    fill_fg_icv: u16,
    fill_bg_icv: u16,
}

/// Workbook-global FONT / XF allocator shared by all sheets.
///
/// Java mapping: POI `HSSFWorkbook` font/style tables. Built-in XF 0..15 are
/// style XFs; 16/17 are date/datetime helpers; custom cell XFs start at
/// [`XF_CUSTOM_BASE`] (18).
#[derive(Debug, Clone, Default)]
pub struct Biff8StyleTable {
    /// Custom fonts beyond the five default Arial records.
    fonts: Vec<FontKey>,
    font_cache: HashMap<FontKey, u16>,
    /// Custom cell XF payloads (indices `XF_CUSTOM_BASE..`).
    xfs: Vec<[u8; 20]>,
    xf_cache: HashMap<XfKey, u16>,
    /// RGB colours allocated into the customizable palette (indices 8..).
    palette_rgb: Vec<(u8, u8, u8)>,
}

impl Biff8StyleTable {
    /// Resolves an XF index for `request`, preserving `base_xf` number format
    /// (`XF_GENERAL` / `XF_DATE` / `XF_DATETIME`).
    pub fn resolve_xf(&mut self, request: &Biff8StyleRequest, base_xf: u16) -> u16 {
        let ifmt = match base_xf {
            XF_DATE => 14,
            XF_DATETIME => 22,
            _ => 0,
        };
        if request.is_default() {
            return base_xf;
        }
        let font_index = self.ensure_font(request);
        let key = XfKey {
            font_index,
            ifmt,
            halign: request.halign.unwrap_or(0),
            valign: request.valign.unwrap_or(2),
            wrap: request.wrap,
            fill_pattern: request.fill_pattern.unwrap_or(0),
            fill_fg_icv: request.fill_fg_icv.unwrap_or(0x40),
            fill_bg_icv: request.fill_bg_icv.unwrap_or(ICV_PATTERN_BG_DEFAULT),
        };
        if let Some(existing) = self.xf_cache.get(&key) {
            return *existing;
        }
        let packed = pack_cell_xf(
            key.font_index,
            key.ifmt,
            key.halign,
            key.valign,
            key.wrap,
            key.fill_pattern,
            key.fill_fg_icv,
            key.fill_bg_icv,
        );
        let index = XF_CUSTOM_BASE + self.xfs.len() as u16;
        self.xfs.push(packed);
        self.xf_cache.insert(key, index);
        index
    }

    /// FONT records after the five defaults (emission order).
    #[must_use]
    pub fn custom_fonts(&self) -> Vec<Vec<u8>> {
        self.fonts
            .iter()
            .map(|font| {
                pack_font(
                    font.height_points,
                    font.bold,
                    font.italic,
                    font.strikeout,
                    font.color_icv,
                    &font.name,
                )
            })
            .collect()
    }

    /// Custom cell XF payloads in emission order.
    #[must_use]
    pub fn custom_xfs(&self) -> &[[u8; 20]] {
        &self.xfs
    }

    /// Whether a PALETTE record is required for custom RGB colours.
    #[must_use]
    pub fn needs_palette(&self) -> bool {
        !self.palette_rgb.is_empty()
    }

    /// Custom RGB colours keyed by palette index starting at 8.
    #[must_use]
    pub fn palette_overrides(&self) -> &[(u8, u8, u8)] {
        &self.palette_rgb
    }

    fn ensure_font(&mut self, request: &Biff8StyleRequest) -> u16 {
        let key = FontKey {
            height_points: request.font_height_points.unwrap_or(10),
            bold: request.bold,
            italic: request.italic,
            strikeout: request.strikeout,
            color_icv: request.font_color_icv.unwrap_or(ICV_AUTO),
            name: request
                .font_name
                .clone()
                .unwrap_or_else(|| "Arial".to_owned()),
        };
        // Default Arial 10 / not bold / auto colour → built-in font 0.
        if key.height_points == 10
            && !key.bold
            && !key.italic
            && !key.strikeout
            && key.color_icv == ICV_AUTO
            && key.name == "Arial"
        {
            return 0;
        }
        if let Some(existing) = self.font_cache.get(&key) {
            return *existing;
        }
        // BIFF8 skips font index 4: slots 0..3 → indices 0..3, slot 4 → index 5, …
        let slot = 5 + self.fonts.len(); // 5th default is index 5; first custom → 6
        let index = font_index_for_slot(slot);
        self.fonts.push(key.clone());
        self.font_cache.insert(key, index);
        index
    }

    /// Allocates or reuses a palette ICV for an RGB triple.
    pub fn alloc_rgb_icv(&mut self, rgb: u32) -> u16 {
        let r = ((rgb >> 16) & 0xFF) as u8;
        let g = ((rgb >> 8) & 0xFF) as u8;
        let b = (rgb & 0xFF) as u8;
        if let Some(pos) = self.palette_rgb.iter().position(|&c| c == (r, g, b)) {
            return (8 + pos) as u16;
        }
        if self.palette_rgb.len() >= 56 {
            // Fall back to nearest built-in when palette is full.
            return nearest_indexed(r, g, b);
        }
        let index = (8 + self.palette_rgb.len()) as u16;
        self.palette_rgb.push((r, g, b));
        index
    }
}

/// Maps FONT record ordinal (0-based among all FONT records) to XF font index.
///
/// Excel / HSSF skip index 4: records `[0,1,2,3,4]` → indices `[0,1,2,3,5]`.
#[must_use]
pub const fn font_index_for_slot(slot: usize) -> u16 {
    if slot < 4 {
        slot as u16
    } else {
        (slot + 1) as u16
    }
}

fn rgb_or_indexed_to_icv(color: ExcelColor) -> u16 {
    match color {
        ExcelColor::Indexed(64) => ICV_AUTO,
        ExcelColor::Indexed(index) => u16::from(index),
        // RGB is approximated to a standard palette entry here; workbook-level
        // custom palette allocation happens via [`Biff8StyleTable::alloc_rgb_icv`]
        // when callers pass through the table. For annotation Indexed colours
        // this path is unused.
        ExcelColor::Rgb(rgb) => nearest_indexed(
            ((rgb >> 16) & 0xFF) as u8,
            ((rgb >> 8) & 0xFF) as u8,
            (rgb & 0xFF) as u8,
        ),
    }
}

/// Converts [`ExcelColor`] using the style table for RGB palette allocation.
#[allow(dead_code)]
pub fn color_to_icv(table: &mut Biff8StyleTable, color: ExcelColor) -> u16 {
    match color {
        ExcelColor::Indexed(64) => ICV_AUTO,
        ExcelColor::Indexed(index) => u16::from(index),
        ExcelColor::Rgb(rgb) => table.alloc_rgb_icv(rgb),
    }
}

fn nearest_indexed(r: u8, g: u8, b: u8) -> u16 {
    // Minimal subset of POI IndexedColors used by Style / Annotation tests.
    const TABLE: &[(u8, u8, u8, u16)] = &[
        (0, 0, 0, 8),
        (255, 255, 255, 9),
        (255, 0, 0, 10),
        (0, 255, 0, 11),
        (0, 0, 255, 12),
        (255, 255, 0, 13),
        (255, 0, 255, 14),
        (0, 255, 255, 15),
        (128, 0, 0, 16),
        (0, 128, 0, 17),
        (0, 0, 128, 18),
        (128, 128, 0, 19),
        (128, 0, 128, 20),
        (0, 128, 128, 21),
        (192, 192, 192, 22),
        (128, 128, 128, 23),
    ];
    let mut best = 8u16;
    let mut best_dist = u32::MAX;
    for &(tr, tg, tb, idx) in TABLE {
        let dr = i32::from(r) - i32::from(tr);
        let dg = i32::from(g) - i32::from(tg);
        let db = i32::from(b) - i32::from(tb);
        let dist = (dr * dr + dg * dg + db * db) as u32;
        if dist < best_dist {
            best_dist = dist;
            best = idx;
        }
    }
    best
}

fn excel_fill_pattern(pattern: ExcelFillPattern) -> u8 {
    // POI `FillPatternType` ordinals.
    match pattern {
        ExcelFillPattern::None => 0,
        ExcelFillPattern::Solid => 1,
        ExcelFillPattern::MediumGray => 2,
        ExcelFillPattern::DarkGray => 3,
        ExcelFillPattern::LightGray => 4,
        ExcelFillPattern::DarkHorizontal => 5,
        ExcelFillPattern::DarkVertical => 6,
        ExcelFillPattern::DarkDown => 7,
        ExcelFillPattern::DarkUp => 8,
        ExcelFillPattern::DarkGrid => 9,
        ExcelFillPattern::DarkTrellis => 10,
        ExcelFillPattern::LightHorizontal => 11,
        ExcelFillPattern::LightVertical => 12,
        ExcelFillPattern::LightDown => 13,
        ExcelFillPattern::LightUp => 14,
        ExcelFillPattern::LightGrid => 15,
        ExcelFillPattern::LightTrellis => 16,
        ExcelFillPattern::Gray125 => 17,
        ExcelFillPattern::Gray0625 => 18,
    }
}

fn excel_halign(align: ExcelHorizontalAlignment) -> u8 {
    match align {
        ExcelHorizontalAlignment::General => 0,
        ExcelHorizontalAlignment::Left => 1,
        ExcelHorizontalAlignment::Center => 2,
        ExcelHorizontalAlignment::Right => 3,
        ExcelHorizontalAlignment::Fill => 4,
        ExcelHorizontalAlignment::Justify => 5,
        ExcelHorizontalAlignment::CenterAcross => 6,
        ExcelHorizontalAlignment::Distributed => 7,
    }
}

fn excel_valign(align: ExcelVerticalAlignment) -> u8 {
    match align {
        ExcelVerticalAlignment::Top => 0,
        ExcelVerticalAlignment::Center => 1,
        ExcelVerticalAlignment::Bottom => 2,
        ExcelVerticalAlignment::Justify => 3,
        ExcelVerticalAlignment::Distributed => 4,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolve_default_keeps_general_xf() {
        let mut table = Biff8StyleTable::default();
        assert_eq!(
            table.resolve_xf(&Biff8StyleRequest::default(), XF_GENERAL),
            XF_GENERAL
        );
    }

    #[test]
    fn resolve_bold_allocates_custom_xf() {
        let mut table = Biff8StyleTable::default();
        let mut req = Biff8StyleRequest::default();
        req.bold = true;
        let xf = table.resolve_xf(&req, XF_GENERAL);
        assert!(xf >= XF_CUSTOM_BASE);
        assert_eq!(table.custom_xfs().len(), 1);
        assert_eq!(table.custom_fonts().len(), 1);
    }

    #[test]
    fn font_index_skips_four() {
        assert_eq!(font_index_for_slot(0), 0);
        assert_eq!(font_index_for_slot(4), 5);
        assert_eq!(font_index_for_slot(5), 6);
    }
}
