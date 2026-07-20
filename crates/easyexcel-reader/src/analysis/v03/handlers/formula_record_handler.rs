//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.FormulaRecordHandler`.
//!
//! Formula string parsing (`HSSFFormulaParser`) stays outside this handler;
//! callers pass the already-resolved formula text and cached result type.

use super::super::xls_record_handler::XlsRecordHandler;

/// Cached formula result kinds aligned with POI `CellType.forInt`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaCachedType {
    /// String result — Java waits for the next `StringRecord`.
    String,
    /// Numeric result.
    Numeric,
    /// Error result (`#VALUE!`).
    Error,
    /// Boolean result.
    Boolean,
    /// Empty / unknown.
    Empty,
}

/// Decoded formula cell produced by [`FormulaRecordHandler`].
#[derive(Debug, Clone, PartialEq)]
pub struct FormulaCell {
    /// Zero-based row.
    pub row: u32,
    /// Zero-based column.
    pub column: usize,
    /// Formula text (may be `None` when parsing failed).
    pub formula: Option<String>,
    /// Cached result type.
    pub cached_type: FormulaCachedType,
    /// Numeric cached value when `cached_type == Numeric`.
    pub number_value: Option<f64>,
    /// Boolean cached value when `cached_type == Boolean`.
    pub bool_value: Option<bool>,
    /// String cached value (`StringRecord` or `#VALUE!` for errors).
    pub string_value: Option<String>,
    /// Whether the string result is pending a following `StringRecord`.
    pub pending_string: bool,
}

/// Mirrors Java `FormulaRecordHandler`.
#[derive(Debug, Default)]
pub struct FormulaRecordHandler {
    /// Pending string-formula cell awaiting `StringRecord`. (Java `tempCellData`)
    pub pending: Option<FormulaCell>,
}

impl FormulaRecordHandler {
    /// Creates an idle handler.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Java `FormulaRecordHandler.processRecord` (caller supplies parsed fields).
    pub fn process_formula(
        &mut self,
        row: u32,
        column: usize,
        formula: Option<String>,
        cached_type: FormulaCachedType,
        number_value: Option<f64>,
        bool_value: Option<bool>,
    ) -> Option<FormulaCell> {
        let string_value = match cached_type {
            FormulaCachedType::Error => Some("#VALUE!".to_owned()),
            _ => None,
        };
        let mut cell = FormulaCell {
            row,
            column,
            formula,
            cached_type,
            number_value,
            bool_value,
            string_value,
            pending_string: false,
        };
        match cached_type {
            FormulaCachedType::String => {
                cell.pending_string = true;
                self.pending = Some(cell);
                None
            }
            FormulaCachedType::Error
            | FormulaCachedType::Numeric
            | FormulaCachedType::Boolean
            | FormulaCachedType::Empty => Some(cell),
        }
    }

    /// Java `StringRecordHandler` follow-up for a pending string formula.
    pub fn complete_pending_string(&mut self, value: String) -> Option<FormulaCell> {
        let mut cell = self.pending.take()?;
        cell.pending_string = false;
        cell.string_value = Some(value);
        Some(cell)
    }
}

/// BIFF `Formula` record sid. (POI `FormulaRecord.sid`)
pub const FORMULA_SID: u16 = 0x0006;

impl XlsRecordHandler for FormulaRecordHandler {
    /// Java `FormulaRecordHandler.processRecord` — parses row/col; full formula
    /// text needs `HSSFFormulaParser` (use [`Self::process_formula`]).
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != FORMULA_SID || data.len() < 4 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        let _ = self.process_formula(row, column, None, FormulaCachedType::Empty, None, None);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn string_formula_is_pending() {
        let mut handler = FormulaRecordHandler::new();
        let immediate = handler.process_formula(
            0,
            0,
            Some("A1".into()),
            FormulaCachedType::String,
            None,
            None,
        );
        assert!(immediate.is_none());
        assert!(handler.pending.is_some());
    }
}
