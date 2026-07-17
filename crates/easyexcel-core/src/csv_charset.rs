//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvCharset` (Java uses `java.nio.charset.Charset`).
//!
//! The Rust port uses a `String` for the Java-style charset label because
//! `encoding_rs::Encoding` (the actual codec backend) is queried separately.

/// Character encoding used by the CSV reader and writer.
///
/// Names follow Java's `Charset.forName` convention. The backend accepts
/// case-insensitive WHATWG labels such as `UTF-8`, `UTF-16BE`, `GBK`, and
/// `windows-1252`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CsvCharset(String);

impl CsvCharset {
    /// Creates a charset from a Java-style charset name or alias.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self(name.into())
    }

    /// Returns the configured charset name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.0
    }

    /// Returns UTF-8, the deterministic Rust default.
    #[must_use]
    pub fn utf8() -> Self {
        Self("UTF-8".to_owned())
    }
}

impl Default for CsvCharset {
    fn default() -> Self {
        Self::utf8()
    }
}

impl From<&str> for CsvCharset {
    fn from(value: &str) -> Self {
        Self::new(value)
    }
}

impl From<String> for CsvCharset {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}
