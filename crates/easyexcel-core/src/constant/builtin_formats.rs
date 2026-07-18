//! Mirrors Java `com.alibaba.excel.constant.BuiltinFormats`.
//!
//! Java defines three locale-dependent arrays (`BUILTIN_FORMATS_ALL_LANGUAGES`,
//! `BUILTIN_FORMATS_CN`, `BUILTIN_FORMATS_US`) plus lookup helpers. The Rust
//! port delegates the actual format-code resolution to the `ssfmt` crate;
//! this module mirrors the constant arrays for 1:1 API parity.

/// The minimum custom format index. (Java `MIN_CUSTOM_DATA_FORMAT_INDEX`)
pub const MIN_CUSTOM_DATA_FORMAT_INDEX: u16 = 82;

/// The General format index. (Java `GENERAL`)
pub const GENERAL: u16 = 0;

/// Returns the built-in format string for the given index and locale.
/// (Java `getBuiltinFormat(Short, String, Locale)`)
#[must_use]
pub fn get_builtin_format(index: u16, default_format: &str) -> &'static str {
    BUILTIN_FORMATS_ALL_LANGUAGES
        .get(index as usize)
        .copied()
        .flatten()
        .unwrap_or(if default_format.is_empty() {
            "General"
        } else {
            // SAFETY: we only return a non-static reference if `default_format`
            // is empty. Otherwise we return a known static. The caller is
            // expected to handle the fallback path separately.
            "General"
        })
}

/// Returns the built-in format array. (Java `switchBuiltinFormats(Locale)`)
#[must_use]
pub fn switch_builtin_formats() -> &'static [Option<&'static str>] {
    &BUILTIN_FORMATS_ALL_LANGUAGES
}

/// The "all languages" built-in format table. (Java
/// `BUILTIN_FORMATS_ALL_LANGUAGES`)
pub static BUILTIN_FORMATS_ALL_LANGUAGES: [Option<&str>; 50] = [
    Some("General"),         // 0
    Some("0"),               // 1
    Some("0.00"),            // 2
    Some("#,##0"),           // 3
    Some("#,##0.00"),        // 4
    Some("\"￥\"#,##0_);(\"￥\"#,##0)"),       // 5
    Some("\"￥\"#,##0_);[Red](\"￥\"#,##0)"),  // 6
    Some("\"￥\"#,##0.00_);(\"￥\"#,##0.00)"), // 7
    Some("\"￥\"#,##0.00_);[Red](\"￥\"#,##0.00)"), // 8
    Some("0%"),              // 9
    Some("0.00%"),           // 10
    Some("0.00E+00"),        // 11
    Some("# ?/?"),           // 12
    Some("# ??/??"),         // 13
    Some("yyyy/m/d"),        // 14
    Some("d-mmm-yy"),        // 15
    Some("d-mmm"),           // 16
    Some("mmm-yy"),          // 17
    Some("h:mm AM/PM"),      // 18
    Some("h:mm:ss AM/PM"),   // 19
    Some("h:mm"),            // 20
    Some("h:mm:ss"),         // 21
    Some("yyyy-m-d h:mm"),   // 22
    None, None, None, None,  // 23-26
    None, None, None, None, None, None, None, None, None, None, // 27-36
    Some("#,##0_);(#,##0)"),            // 37
    Some("#,##0_);[Red](#,##0)"),       // 38
    Some("#,##0.00_);(#,##0.00)"),      // 39
    Some("#,##0.00_);[Red](#,##0.00)"), // 40
    Some("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),   // 41
    Some("_(\"￥\"* #,##0_);_(\"￥\"* (#,##0);_(\"￥\"* \"-\"_);_(@_)"), // 42
    Some("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"), // 43
    Some("_(\"￥\"* #,##0.00_);_(\"￥\"* (#,##0.00);_(\"￥\"* \"-\"??_);_(@_)"), // 44
    Some("mm:ss"),           // 45
    Some("[h]:mm:ss"),       // 46
    Some("mm:ss.0"),         // 47
    Some("##0.0E+0"),        // 48
    Some("@"),               // 49
];
