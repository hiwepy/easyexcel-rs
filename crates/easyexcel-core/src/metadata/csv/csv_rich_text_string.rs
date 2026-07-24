//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvRichTextString`.

/// CSV rich text wrapper.
///
/// CSV cannot preserve font runs, so Java stores only the plain string and
/// makes formatting methods inert. Rust exposes the same meaningful state.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CsvRichTextString {
    value: String,
}

impl CsvRichTextString {
    /// Creates a CSV rich-text value from plain text.
    #[must_use]
    pub fn new(value: impl Into<String>) -> Self {
        Self {
            value: value.into(),
        }
    }

    /// Returns the plain text. (Java `getString()`)
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.value
    }

    /// Returns the UTF-8 text length in Unicode scalar values.
    #[must_use]
    pub fn len(&self) -> usize {
        self.value.chars().count()
    }

    /// Returns whether the text is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.value.is_empty()
    }
}
