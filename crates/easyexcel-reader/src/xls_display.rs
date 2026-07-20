//! XLS BIFF format overlay for STRING-mode display.
//!
//! Calamine converts date-formatted numbers to `Data::DateTime` and discards the
//! Excel format code. EasyExcel / POI keep the XF → FORMAT mapping and render
//! via `DataFormatter` (BuiltinFormats + custom codes). This module re-reads the
//! Workbook stream's FORMAT / XF / NUMBER / RK / MulRk records so Rust STRING
//! mode can match Java short dates (`yyyy-m-d h:mm`) and related formats.

use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

use cfb::CompoundFile;
use easyexcel_core::constant::builtin_format_code;
use ssfmt::Locale;

use crate::xlsx_rows::format_with_code;

/// Per-sheet map of `(row, col) → formatted STRING display`.
pub(crate) type SheetDisplays = HashMap<(u32, usize), String>;

/// Load formatted display strings for every numeric cell in an `.xls` workbook.
///
/// Returns one map per worksheet (BoundSheet order). Failures are soft: callers
/// may fall back to calamine `as_text()` when a sheet map is missing.
pub(crate) fn load_xls_displays(path: &Path, date_1904: bool, locale: &Locale) -> Vec<SheetDisplays> {
    match load_xls_displays_inner(path, date_1904, locale) {
        Ok(sheets) => sheets,
        Err(_) => Vec::new(),
    }
}

fn load_xls_displays_inner(
    path: &Path,
    date_1904: bool,
    locale: &Locale,
) -> Result<Vec<SheetDisplays>, String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let mut cfb = CompoundFile::open(file).map_err(|e| e.to_string())?;
    let mut stream = cfb
        .open_stream("/Workbook")
        .or_else(|_| cfb.open_stream("/Book"))
        .map_err(|e| e.to_string())?;
    let mut wb = Vec::new();
    stream.read_to_end(&mut wb).map_err(|e| e.to_string())?;
    Ok(parse_workbook_displays(&wb, date_1904, locale))
}

fn parse_workbook_displays(wb: &[u8], date_1904: bool, locale: &Locale) -> Vec<SheetDisplays> {
    let mut custom_formats: HashMap<u16, String> = HashMap::new();
    let mut xfs: Vec<u16> = Vec::new();
    let mut sheets: Vec<SheetDisplays> = Vec::new();
    let mut sheet_idx: isize = -1;
    let mut in_sheet = false;
    let mut i = 0usize;

    while i + 4 <= wb.len() {
        let typ = u16::from_le_bytes([wb[i], wb[i + 1]]);
        let length = u16::from_le_bytes([wb[i + 2], wb[i + 3]]) as usize;
        i += 4;
        if i + length > wb.len() {
            break;
        }
        let payload = &wb[i..i + length];
        i += length;

        match typ {
            0x0809 if payload.len() >= 4 => {
                let dt = u16::from_le_bytes([payload[2], payload[3]]);
                if dt == 0x0010 {
                    sheet_idx += 1;
                    while sheets.len() as isize <= sheet_idx {
                        sheets.push(HashMap::new());
                    }
                    in_sheet = true;
                } else {
                    in_sheet = false;
                }
            }
            0x041E => {
                if let Some((ifmt, code)) = parse_format_record(payload) {
                    custom_formats.insert(ifmt, code);
                }
            }
            0x00E0 if payload.len() >= 4 => {
                let ifmt = u16::from_le_bytes([payload[2], payload[3]]);
                xfs.push(ifmt);
            }
            0x0203 if in_sheet && payload.len() >= 14 => {
                let row = u16::from_le_bytes([payload[0], payload[1]]) as u32;
                let col = u16::from_le_bytes([payload[2], payload[3]]) as usize;
                let xf = u16::from_le_bytes([payload[4], payload[5]]) as usize;
                let value = f64::from_le_bytes(payload[6..14].try_into().unwrap_or([0; 8]));
                push_display(
                    &mut sheets,
                    sheet_idx,
                    row,
                    col,
                    xf,
                    value,
                    &xfs,
                    &custom_formats,
                    date_1904,
                    locale,
                );
            }
            0x027E if in_sheet && payload.len() >= 10 => {
                let row = u16::from_le_bytes([payload[0], payload[1]]) as u32;
                let col = u16::from_le_bytes([payload[2], payload[3]]) as usize;
                let xf = u16::from_le_bytes([payload[4], payload[5]]) as usize;
                let value = decode_rk(&payload[6..10]);
                push_display(
                    &mut sheets,
                    sheet_idx,
                    row,
                    col,
                    xf,
                    value,
                    &xfs,
                    &custom_formats,
                    date_1904,
                    locale,
                );
            }
            0x00BD if in_sheet && payload.len() >= 6 => {
                // MulRk: row, firstCol, then repeating (xf, rk) until lastCol
                let row = u16::from_le_bytes([payload[0], payload[1]]) as u32;
                let first_col = u16::from_le_bytes([payload[2], payload[3]]) as usize;
                let last_col = u16::from_le_bytes([
                    payload[payload.len() - 2],
                    payload[payload.len() - 1],
                ]) as usize;
                let mut offset = 4usize;
                let mut col = first_col;
                while col <= last_col && offset + 6 <= payload.len().saturating_sub(2) {
                    let xf = u16::from_le_bytes([payload[offset], payload[offset + 1]]) as usize;
                    let value = decode_rk(&payload[offset + 2..offset + 6]);
                    push_display(
                        &mut sheets,
                        sheet_idx,
                        row,
                        col,
                        xf,
                        value,
                        &xfs,
                        &custom_formats,
                        date_1904,
                        locale,
                    );
                    offset += 6;
                    col += 1;
                }
            }
            _ => {}
        }
    }
    sheets
}

fn push_display(
    sheets: &mut [SheetDisplays],
    sheet_idx: isize,
    row: u32,
    col: usize,
    xf: usize,
    value: f64,
    xfs: &[u16],
    custom_formats: &HashMap<u16, String>,
    date_1904: bool,
    locale: &Locale,
) {
    if sheet_idx < 0 || !value.is_finite() {
        return;
    }
    let Some(ifmt) = xfs.get(xf).copied() else {
        return;
    };
    let code = custom_formats
        .get(&ifmt)
        .map(String::as_str)
        .or_else(|| builtin_format_code(ifmt));
    let Some(code) = code else {
        return;
    };
    // General / @ — leave to calamine textualization.
    if code.eq_ignore_ascii_case("General") || code == "@" {
        return;
    }
    let Some(display) = format_with_code(value, code, date_1904, locale) else {
        return;
    };
    if let Some(sheet) = sheets.get_mut(sheet_idx as usize) {
        sheet.insert((row, col), display);
    }
}

fn parse_format_record(payload: &[u8]) -> Option<(u16, String)> {
    if payload.len() < 5 {
        return None;
    }
    let ifmt = u16::from_le_bytes([payload[0], payload[1]]);
    let cch = u16::from_le_bytes([payload[2], payload[3]]) as usize;
    let flags = payload[4];
    let raw = &payload[5..];
    let code = if flags & 1 != 0 {
        let bytes = cch.saturating_mul(2).min(raw.len());
        if bytes < 2 {
            return None;
        }
        let units: Vec<u16> = raw[..bytes]
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        String::from_utf16_lossy(&units)
    } else {
        // BIFF8 compressed Unicode is Latin-1 code units (one byte per char),
        // not UTF-8. Byte `0xA5` must stay `¥` (U+00A5); UTF-8 lossy decode
        // would turn it into U+FFFD and break STRING currency cells.
        let bytes = cch.min(raw.len());
        raw[..bytes].iter().map(|&b| b as char).collect()
    };
    Some((ifmt, code))
}

/// Decode an RK number (see MS-XLS 2.5.209).
fn decode_rk(bytes: &[u8]) -> f64 {
    if bytes.len() < 4 {
        return 0.0;
    }
    let rk = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let d100 = rk & 0x01 != 0;
    let is_int = rk & 0x02 != 0;
    let value = if is_int {
        ((rk as i32) >> 2) as f64
    } else {
        f64::from_bits((u64::from(rk & !0x03)) << 32)
    };
    if d100 { value / 100.0 } else { value }
}

#[cfg(test)]
mod tests {
    use super::*;
    use ssfmt::Locale;

    #[test]
    fn decode_rk_integer() {
        // 100 as integer RK (bit1 set), little-endian packed
        let rk = (100i32 << 2) as u32 | 0x02;
        let bytes = rk.to_le_bytes();
        assert!((decode_rk(&bytes) - 100.0).abs() < f64::EPSILON);
    }

    #[test]
    fn java_compat_percent_via_format_with_code() {
        let locale = Locale::default();
        assert_eq!(
            format_with_code(99.99, "#.##%", false, &locale).as_deref(),
            Some("9999%")
        );
    }

    /// POI / EasyExcel: `_ ` pads are dropped; `\ ` on the negative section is kept.
    #[test]
    fn java_compat_trailing_space_accounting_format() {
        let locale = Locale::default();
        let code = r"0.00_ ;[Red]\-0.00\ ";
        assert_eq!(
            format_with_code(-1.07, code, false, &locale).as_deref(),
            Some("-1.07 ")
        );
        assert_eq!(
            format_with_code(14.11, code, false, &locale).as_deref(),
            Some("14.11")
        );
        // Accounting `_)` must not leave a trailing pad space (Java `24.20`).
        let acct = r"0.00_);[Red]\(0.00\)";
        assert_eq!(
            format_with_code(24.199812400000013, acct, false, &locale).as_deref(),
            Some("24.20")
        );
    }

    /// DateFormatTest#t03Read — unpadded month `yyyy-m-dd` → `2023-1-01`.
    #[test]
    fn java_compat_short_month_dataformat_v2() {
        let locale = Locale::default();
        let code = r"yyyy\-m\-dd\ hh:mm:ss";
        // Excel serial for 2023-01-01
        assert_eq!(
            format_with_code(44927.0, code, false, &locale).as_deref(),
            Some("2023-1-01 00:00:00")
        );
    }

    /// CN `上午/下午` literal must resolve via locale AM/PM (not printed as slash text).
    #[test]
    fn java_compat_cn_ampm_literal_resolves() {
        let locale = Locale {
            decimal_separator: '.',
            thousands_separator: ',',
            currency_symbol: "¥",
            am_string: "上午",
            pm_string: "下午",
            month_names_short: [
                "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月",
                "12月",
            ],
            month_names_full: [
                "一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月",
                "十一月", "十二月",
            ],
            day_names_short: ["日", "一", "二", "三", "四", "五", "六"],
            day_names_full: [
                "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六",
            ],
        };
        // 2020-01-01 01:01
        let serial = 43831.0 + 1.0 / 24.0 + 1.0 / 1440.0;
        assert_eq!(
            format_with_code(serial, r#"[DBNum1]上午/下午h"时"mm"分""#, false, &locale)
                .as_deref(),
            Some("上午1时01分")
        );
        assert_eq!(
            format_with_code(serial, "mmmmm/yy", false, &locale).as_deref(),
            Some("\u{E001}1月\u{E002}/20")
        );
    }

    /// BIFF8 compressed Unicode FORMAT records are Latin-1 (¥ = 0xA5), not UTF-8.
    #[test]
    fn parse_format_record_latin1_yen() {
        // ifmt=5, cch=len, flags=0, body = `"¥"#,##0` in Latin-1
        let mut payload = vec![5, 0, 0, 0, 0];
        let body: Vec<u8> = b"\"\xA5\"#,##0".to_vec();
        payload[2] = body.len() as u8;
        payload.extend_from_slice(&body);
        let (ifmt, code) = parse_format_record(&payload).expect("FORMAT");
        assert_eq!(ifmt, 5);
        assert_eq!(code, "\"¥\"#,##0");
        assert!(!code.contains('\u{FFFD}'));
    }
}
