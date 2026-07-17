//! Mirrors Java `com.alibaba.excel.enums.ByteOrderMarkEnum`.
//!
//! Maps CSV charset names to their leading BOM. Java uses
//! `org.apache.commons.io.ByteOrderMark`; Rust uses byte literal arrays.

use crate::ExcelError;

/// UTF BOM byte sequences aligned with Java's `ByteOrderMarkEnum`.
///
/// Rust port of Java `ByteOrderMarkEnum`. Stores the raw BOM bytes and the
/// associated canonical charset name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ByteOrderMark {
    /// UTF-8 BOM (`EF BB BF`).
    Utf8,
    /// UTF-16 big-endian BOM (`FE FF`).
    Utf16Be,
    /// UTF-16 little-endian BOM (`FF FE`).
    Utf16Le,
    /// UTF-32 big-endian BOM (`00 00 FE FF`).
    Utf32Be,
    /// UTF-32 little-endian BOM (`FF FE 00 00`).
    Utf32Le,
}

impl ByteOrderMark {
    /// Returns the BOM bytes as a slice.
    #[must_use]
    pub const fn bytes(self) -> &'static [u8] {
        match self {
            Self::Utf8 => &[0xEF, 0xBB, 0xBF],
            Self::Utf16Be => &[0xFE, 0xFF],
            Self::Utf16Le => &[0xFF, 0xFE],
            Self::Utf32Be => &[0x00, 0x00, 0xFE, 0xFF],
            Self::Utf32Le => &[0xFF, 0xFE, 0x00, 0x00],
        }
    }

    /// Canonical charset name matched against the BOM.
    #[must_use]
    pub const fn charset_name(self) -> &'static str {
        match self {
            Self::Utf8 => "UTF-8",
            Self::Utf16Be => "UTF-16BE",
            Self::Utf16Le => "UTF-16LE",
            Self::Utf32Be => "UTF-32BE",
            Self::Utf32Le => "UTF-32LE",
        }
    }

    /// Resolves a Java-style charset label to its BOM, if any.
    #[must_use]
    pub fn value_of_by_charset_name(name: &str) -> Option<Self> {
        match name.to_ascii_uppercase().as_str() {
            "UTF-8" | "UTF8" => Some(Self::Utf8),
            "UTF-16BE" => Some(Self::Utf16Be),
            "UTF-16LE" => Some(Self::Utf16Le),
            "UTF-32BE" => Some(Self::Utf32Be),
            "UTF-32LE" => Some(Self::Utf32Le),
            _ => None,
        }
    }

    /// Returns an error explaining the BOM lookup failure (for `Result`-style callers).
    #[must_use]
    pub fn error_for_missing_bom(name: &str) -> ExcelError {
        ExcelError::Unsupported(format!("unsupported CSV charset: {name}"))
    }
}
