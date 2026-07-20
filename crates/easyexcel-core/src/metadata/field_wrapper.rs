//! Mirrors Java `com.alibaba.excel.metadata.FieldWrapper`.

/// Runtime field metadata for one annotated model field.
///
/// Java stores a reflective `Field`. Rust stores the field name and header
/// labels because `#[derive(ExcelRow)]` resolves reflection at compile time.
///
/// Rust port of Java `FieldWrapper`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FieldWrapper {
    /// Rust field name. (Java `field` / `fieldName`)
    pub field_name: String,
    /// Sheet header labels from `@ExcelProperty`. (Java `heads`)
    pub heads: Vec<String>,
}

impl FieldWrapper {
    /// Creates a field wrapper. (Java all-args constructor)
    #[must_use]
    pub fn new(field_name: impl Into<String>, heads: Vec<String>) -> Self {
        Self {
            field_name: field_name.into(),
            heads,
        }
    }

    /// Returns the field name. (Java `getFieldName()`)
    #[must_use]
    pub fn field_name(&self) -> &str {
        &self.field_name
    }

    /// Returns the configured header labels. (Java `getHeads()`)
    #[must_use]
    pub fn heads(&self) -> &[String] {
        &self.heads
    }
}
