//! Mirrors Java `com.alibaba.excel.metadata.property.ExcelDataValidationProperty`
//! (introduced in Phase 1, derived from `@ExcelDataValidation` annotation).

/// Static metadata describing a per-cell data-validation rule.
///
/// Mirrors the subset of fields that Java exposes through
/// `ExcelDataValidationProperty` / `com.alibaba.excel.annotation.write.ExcelDataValidation`.
///
/// Default-constructible + `Copy` so it can sit inside `ExcelColumn`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ExcelDataValidationMeta {
    /// Validation type. (Java `@ExcelDataValidation.type()`, e.g. "list", "whole", "decimal")
    pub data_type: &'static str,
    /// Comparison operator. (Java `@ExcelDataValidation.operator()`, e.g. "between")
    pub operator: &'static str,
    /// First operand expression. (Java `@ExcelDataValidation.formula1()`)
    pub formula1: &'static str,
    /// Second operand expression. (Java `@ExcelDataValidation.formula2()`)
    pub formula2: &'static str,
}

impl ExcelDataValidationMeta {
    /// Builds a new metadata record. (Mirrors Java default constructor + setters)
    #[must_use]
    pub const fn new(
        data_type: &'static str,
        operator: &'static str,
        formula1: &'static str,
        formula2: &'static str,
    ) -> Self {
        Self {
            data_type,
            operator,
            formula1,
            formula2,
        }
    }

    /// Whether this rule is non-empty.
    #[must_use]
    pub const fn is_present(&self) -> bool {
        !self.data_type.is_empty()
    }
}
