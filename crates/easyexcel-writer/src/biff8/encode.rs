//! Low-level BIFF8 record framing and XLUnicodeString helpers.
//!
//! Record layout matches [MS-XLS] / OpenOffice BIFF8: 2-byte type, 2-byte length,
//! then payload (≤ 8224 bytes). Unicode strings use the short / long XLUnicode
//! forms that calamine and Excel expect.
//!
//! XF / COLINFO / ROW / MERGECELLS layouts follow the same field packing as
//! xlwt / Apache POI HSSF (see module docs on each writer helper).

#![allow(dead_code)]

/// Maximum BIFF record data payload (excluding the 4-byte header).
pub const MAX_RECORD_DATA: usize = 8224;

pub const BOF: u16 = 0x0809;
pub const EOF: u16 = 0x000A;
pub const BOUNDSHEET: u16 = 0x0085;
pub const SST: u16 = 0x00FC;
pub const EXTSST: u16 = 0x00FF;
pub const LABELSST: u16 = 0x00FD;
/// Inline Unicode string cell (BIFF8 `LABEL` 0x0204) — used by template writes
/// that must not mutate the shared SST.
pub const LABEL: u16 = 0x0204;
pub const NUMBER: u16 = 0x0203;
pub const RK: u16 = 0x027E;
pub const BOOLERR: u16 = 0x0205;
pub const BLANK: u16 = 0x0201;
pub const XF: u16 = 0x00E0;
pub const FORMAT: u16 = 0x041E;
pub const FONT: u16 = 0x0031;
pub const DIMENSION: u16 = 0x0200;
pub const DATEMODE: u16 = 0x0022;
pub const CODEPAGE: u16 = 0x0042;
pub const CONTINUE: u16 = 0x003C;
pub const WINDOW2: u16 = 0x023E;
pub const STYLE: u16 = 0x0293;
/// Column width / default XF for a column range. (Java HSSF `ColumnInfoRecord`)
pub const COLINFO: u16 = 0x007D;
/// Per-row height and outline flags. (Java HSSF `RowRecord`)
pub const ROW: u16 = 0x0208;
/// Merged-cell ranges. (Java HSSF `MergedCellsTable` / record 0x00E5)
pub const MERGECELLS: u16 = 0x00E5;
/// Custom palette colours (indices 8..). (Java HSSF `PaletteRecord`)
pub const PALETTE: u16 = 0x0092;

/// Workbook globals substream type.
pub const DT_GLOBALS: u16 = 0x0005;
/// Worksheet substream type.
pub const DT_WORKSHEET: u16 = 0x0010;
/// BIFF8 version word.
pub const BIFF8_VERSION: u16 = 0x0600;

/// Built-in XF index used for unstyled cells (last of the 16 style XFs).
pub const XF_GENERAL: u16 = 15;
/// First cell XF after the 16 built-in style XFs — date (`m/d/yy`, id 14).
pub const XF_DATE: u16 = 16;
/// Second cell XF — datetime (`m/d/yy h:mm`, id 22).
pub const XF_DATETIME: u16 = 17;
/// First custom cell XF index (after date / datetime helpers).
pub const XF_CUSTOM_BASE: u16 = 18;

/// Automatic / default font colour ICV.
pub const ICV_AUTO: u16 = 0x7FFF;
/// Default pattern background (automatic).
pub const ICV_PATTERN_BG_DEFAULT: u16 = 64;

/// Appends a framed BIFF record (`type` + `len` + `data`) to `out`.
pub fn record(out: &mut Vec<u8>, typ: u16, data: &[u8]) {
    debug_assert!(data.len() <= MAX_RECORD_DATA);
    out.extend_from_slice(&typ.to_le_bytes());
    out.extend_from_slice(&(data.len() as u16).to_le_bytes());
    out.extend_from_slice(data);
}

/// Encodes a long XLUnicodeString (`cch:u16` + `grbit` + chars).
pub fn encode_unicode_string(s: &str) -> Vec<u8> {
    let chars: Vec<u16> = s.encode_utf16().collect();
    let compressed = chars.iter().all(|&c| c <= 0xFF);
    let mut out = Vec::with_capacity(3 + chars.len() * if compressed { 1 } else { 2 });
    out.extend_from_slice(&(chars.len() as u16).to_le_bytes());
    if compressed {
        out.push(0x00);
        for &c in &chars {
            out.push(c as u8);
        }
    } else {
        out.push(0x01);
        for &c in &chars {
            out.extend_from_slice(&c.to_le_bytes());
        }
    }
    out
}

/// Encodes a short XLUnicodeString (`cch:u8` + `grbit` + chars) for BOUNDSHEET / FONT.
pub fn encode_short_unicode_string(s: &str) -> Vec<u8> {
    let chars: Vec<u16> = s.encode_utf16().take(255).collect();
    let compressed = chars.iter().all(|&c| c <= 0xFF);
    let mut out = Vec::with_capacity(2 + chars.len() * if compressed { 1 } else { 2 });
    out.push(chars.len() as u8);
    if compressed {
        out.push(0x00);
        for &c in &chars {
            out.push(c as u8);
        }
    } else {
        out.push(0x01);
        for &c in &chars {
            out.extend_from_slice(&c.to_le_bytes());
        }
    }
    out
}

/// Tries to pack `v` into an RK record value; `None` means emit a NUMBER record.
pub fn encode_rk(v: f64) -> Option<u32> {
    if !v.is_finite() {
        return None;
    }
    // Integer form (bit0=0, bit1=1).
    if v.fract() == 0.0 && v >= -0x1FFF_FFFF as f64 && v <= 0x1FFF_FFFF as f64 {
        #[allow(clippy::cast_possible_truncation)]
        let n = v as i32;
        return Some(((n as u32) << 2) | 0x02);
    }
    // Integer / 100 form (bit0=1, bit1=1).
    let scaled = v * 100.0;
    if scaled.fract() == 0.0 && scaled >= -0x1FFF_FFFF as f64 && scaled <= 0x1FFF_FFFF as f64 {
        #[allow(clippy::cast_possible_truncation)]
        let n = scaled as i32;
        return Some(((n as u32) << 2) | 0x03);
    }
    // Truncated IEEE754 high 30 bits (bit0=0, bit1=0).
    let bits = v.to_bits();
    let high = (bits >> 32) as u32;
    if (bits & 0xFFFF_FFFF) == 0 && (high & 0x3) == 0 {
        return Some(high);
    }
    None
}

/// Packs a BIFF8 cell XF (20 bytes) with optional solid fill / alignment.
///
/// Field packing matches xlwt `XFRecord` / OpenOffice BIFF8 XF (Java HSSF
/// `ExtendedFormatRecord`).
///
/// `halign` / `valign` use POI codes (`HorizontalAlignment` /
/// `VerticalAlignment` ordinals). `fill_pattern` uses POI `FillPatternType`
/// codes (`SolidForeground = 1`).
#[must_use]
pub fn pack_cell_xf(
    font_index: u16,
    ifmt: u16,
    halign: u8,
    valign: u8,
    wrap: bool,
    fill_pattern: u8,
    fill_fg_icv: u16,
    fill_bg_icv: u16,
) -> [u8; 20] {
    let mut d = [0u8; 20];
    d[0..2].copy_from_slice(&font_index.to_le_bytes());
    d[2..4].copy_from_slice(&ifmt.to_le_bytes());
    // fLocked cell XF (not style).
    d[4..6].copy_from_slice(&0x0001u16.to_le_bytes());
    let mut align = halign & 0x07;
    if wrap {
        align |= 0x08;
    }
    // Vertical alignment in bits 4-6 (default bottom = 2 when unset).
    align |= (valign & 0x07) << 4;
    d[6] = align;
    d[7] = 0; // rotation
    d[8] = 0; // indent / shrink
    d[9] = 0xF8; // XF_USED_ATTRIB — all groups used (cell XF)
    // brd1 (bytes 10-13) left zero — no borders.
    // brd2 (bytes 14-17): fill pattern in bits 26-31.
    let brd2 = u32::from(fill_pattern & 0x3F) << 26;
    d[14..18].copy_from_slice(&brd2.to_le_bytes());
    // pattern colours (bytes 18-19).
    let pat = (fill_fg_icv & 0x7F) | ((fill_bg_icv & 0x7F) << 7);
    d[18..20].copy_from_slice(&(pat as u16).to_le_bytes());
    d
}

/// Packs a BIFF8 FONT record payload (Java HSSF `FontRecord`).
///
/// `height_points` is converted to twips (`* 20`). Bold uses `bls=700`.
#[must_use]
pub fn pack_font(
    height_points: u16,
    bold: bool,
    italic: bool,
    strikeout: bool,
    color_icv: u16,
    name: &str,
) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(height_points.saturating_mul(20)).to_le_bytes());
    let mut grbit = 0u16;
    if italic {
        grbit |= 0x02;
    }
    if strikeout {
        grbit |= 0x08;
    }
    data.extend_from_slice(&grbit.to_le_bytes());
    data.extend_from_slice(&color_icv.to_le_bytes());
    data.extend_from_slice(&(if bold { 700u16 } else { 400u16 }).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes()); // sss
    data.extend_from_slice(&[0, 0, 0, 0]); // uls, family, charset, reserved
    data.extend_from_slice(&encode_short_unicode_string(name));
    data
}

/// Packs one MERGECELLS range (8 bytes): `rwFirst..rwLast`, `colFirst..colLast`.
///
/// Java HSSF `MergedCellsTable` / record 0x00E5.
#[must_use]
pub fn pack_merge_range(first_row: u16, last_row: u16, first_col: u16, last_col: u16) -> [u8; 8] {
    let mut d = [0u8; 8];
    d[0..2].copy_from_slice(&first_row.to_le_bytes());
    d[2..4].copy_from_slice(&last_row.to_le_bytes());
    d[4..6].copy_from_slice(&first_col.to_le_bytes());
    d[6..8].copy_from_slice(&last_col.to_le_bytes());
    d
}

/// Emits one or more MERGECELLS records (max 1027 ranges each).
pub fn write_merge_cells(out: &mut Vec<u8>, ranges: &[[u8; 8]]) {
    const MAX_PER_RECORD: usize = 1027;
    for chunk in ranges.chunks(MAX_PER_RECORD) {
        let mut data = Vec::with_capacity(2 + chunk.len() * 8);
        data.extend_from_slice(&(chunk.len() as u16).to_le_bytes());
        for range in chunk {
            data.extend_from_slice(range);
        }
        record(out, MERGECELLS, &data);
    }
}

/// Writes a PALETTE record with optional RGB overrides at indices 8..
///
/// Java HSSF `PaletteRecord` — first override replaces palette slot 8, etc.
pub fn write_palette_record(out: &mut Vec<u8>, overrides: &[(u8, u8, u8)]) {
    // Standard BIFF8 customizable palette (56 colours, indices 8..63).
    let mut colours: [(u8, u8, u8); 56] = [
        (0, 0, 0),
        (255, 255, 255),
        (255, 0, 0),
        (0, 255, 0),
        (0, 0, 255),
        (255, 255, 0),
        (255, 0, 255),
        (0, 255, 255),
        (128, 0, 0),
        (0, 128, 0),
        (0, 0, 128),
        (128, 128, 0),
        (128, 0, 128),
        (0, 128, 128),
        (192, 192, 192),
        (128, 128, 128),
        (153, 153, 255),
        (153, 51, 102),
        (255, 255, 204),
        (204, 255, 255),
        (102, 0, 102),
        (255, 128, 128),
        (0, 102, 204),
        (204, 204, 255),
        (0, 0, 128),
        (255, 0, 255),
        (255, 255, 0),
        (0, 255, 255),
        (128, 0, 128),
        (128, 0, 0),
        (0, 128, 128),
        (0, 0, 255),
        (0, 204, 255),
        (204, 255, 255),
        (204, 255, 204),
        (255, 255, 153),
        (153, 204, 255),
        (255, 153, 204),
        (204, 153, 255),
        (255, 204, 153),
        (51, 102, 255),
        (51, 204, 204),
        (153, 204, 0),
        (255, 204, 0),
        (255, 153, 0),
        (255, 102, 0),
        (102, 102, 153),
        (150, 150, 150),
        (0, 51, 102),
        (51, 153, 102),
        (0, 51, 0),
        (51, 51, 0),
        (153, 51, 0),
        (153, 51, 102),
        (51, 51, 153),
        (51, 51, 51),
    ];
    for (i, rgb) in overrides.iter().enumerate() {
        if i < colours.len() {
            colours[i] = *rgb;
        }
    }
    let mut data = Vec::with_capacity(2 + colours.len() * 4);
    data.extend_from_slice(&(colours.len() as u16).to_le_bytes());
    for &(r, g, b) in &colours {
        data.push(r);
        data.push(g);
        data.push(b);
        data.push(0);
    }
    record(out, PALETTE, &data);
}

/// Packs a COLINFO record payload (12 bytes).
///
/// `width_chars` is Excel's character width; stored as `width * 256` (POI
/// `sheet.setColumnWidth(col, chars * 256)`).
#[must_use]
pub fn pack_colinfo(first_col: u8, last_col: u8, width_chars: u16, xf_index: u16) -> [u8; 12] {
    let coldx = width_chars.saturating_mul(256);
    let mut d = [0u8; 12];
    d[0..2].copy_from_slice(&u16::from(first_col).to_le_bytes());
    d[2..4].copy_from_slice(&u16::from(last_col).to_le_bytes());
    d[4..6].copy_from_slice(&coldx.to_le_bytes());
    d[6..8].copy_from_slice(&xf_index.to_le_bytes());
    // options + unused remain zero.
    d
}

/// Packs a ROW record payload (16 bytes).
///
/// `height_points` is converted to twips (`* 20`), matching POI
/// `row.setHeightInPoints` / Java StyleDataTest (`40pt → 800`).
#[must_use]
pub fn pack_row(row: u16, first_col: u8, last_col_exclusive: u8, height_points: u16) -> [u8; 16] {
    let miy = height_points.saturating_mul(20) & 0x7FFF; // bit15=0 → custom height
    let mut d = [0u8; 16];
    d[0..2].copy_from_slice(&row.to_le_bytes());
    d[2..4].copy_from_slice(&u16::from(first_col).to_le_bytes());
    d[4..6].copy_from_slice(&u16::from(last_col_exclusive).to_le_bytes());
    d[6..8].copy_from_slice(&miy.to_le_bytes());
    // unused + reserved
    // option flags: bit8 always 1 (0x100); bit6 = height unsynced (0x40)
    let options: u32 = 0x0100 | 0x0040;
    d[12..16].copy_from_slice(&options.to_le_bytes());
    d
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pack_colinfo_matches_poi_character_units() {
        let bytes = pack_colinfo(0, 0, 50, XF_GENERAL);
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 50 * 256);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), XF_GENERAL);
    }

    #[test]
    fn pack_row_matches_poi_twips() {
        let bytes = pack_row(0, 0, 2, 40);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), 800);
    }

    #[test]
    fn pack_cell_xf_solid_yellow() {
        // IndexedColors.YELLOW = 13, solid pattern = 1, valign bottom = 2.
        let bytes = pack_cell_xf(0, 0, 0, 2, false, 1, 13, ICV_PATTERN_BG_DEFAULT);
        let brd2 = u32::from_le_bytes([bytes[14], bytes[15], bytes[16], bytes[17]]);
        assert_eq!((brd2 >> 26) & 0x3F, 1);
        let pat = u16::from_le_bytes([bytes[18], bytes[19]]);
        assert_eq!(pat & 0x7F, 13);
        assert_eq!((pat >> 7) & 0x7F, ICV_PATTERN_BG_DEFAULT);
        assert_eq!((bytes[6] >> 4) & 0x07, 2);
    }

    #[test]
    fn pack_font_bold_arial_12() {
        let bytes = pack_font(12, true, false, false, ICV_AUTO, "Arial");
        assert_eq!(u16::from_le_bytes([bytes[0], bytes[1]]), 240);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), 700);
    }

    #[test]
    fn pack_merge_range_layout() {
        let bytes = pack_merge_range(1, 2, 0, 1);
        assert_eq!(u16::from_le_bytes([bytes[0], bytes[1]]), 1);
        assert_eq!(u16::from_le_bytes([bytes[2], bytes[3]]), 2);
        assert_eq!(u16::from_le_bytes([bytes[4], bytes[5]]), 0);
        assert_eq!(u16::from_le_bytes([bytes[6], bytes[7]]), 1);
    }
}
