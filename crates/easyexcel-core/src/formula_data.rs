//! Mirrors Java `com.alibaba.excel.metadata.data.FormulaData`.

/// Formula metadata associated with a cached cell value while reading.
///
/// Mirrors Java `FormulaData` (`formulaValue` field + `clone()` override).
/// Rust uses `#[derive(Clone)]` so the public `clone()` is automatic.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FormulaData {
    formula_value: String,
}

impl FormulaData {
    /// Creates formula metadata from the expression stored in the workbook.
    #[must_use]
    pub fn new(formula_value: impl Into<String>) -> Self {
        Self {
            formula_value: formula_value.into(),
        }
    }

    /// Returns the formula expression without adding a leading equals sign. (Java `getFormulaValue()`)
    #[must_use]
    pub fn formula_value(&self) -> &str {
        &self.formula_value
    }
}
