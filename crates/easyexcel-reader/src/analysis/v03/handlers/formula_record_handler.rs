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
    /// XF index used by the cached numeric result.
    pub format_index: u16,
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
    /// Most recently completed formula cell.
    pub last_cell: Option<FormulaCell>,
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
        self.process_formula_with_format(
            row,
            column,
            formula,
            cached_type,
            number_value,
            bool_value,
            0,
        )
    }

    /// Java `FormulaRecordHandler.processRecord` with the record XF index.
    pub fn process_formula_with_format(
        &mut self,
        row: u32,
        column: usize,
        formula: Option<String>,
        cached_type: FormulaCachedType,
        number_value: Option<f64>,
        bool_value: Option<bool>,
        format_index: u16,
    ) -> Option<FormulaCell> {
        let string_value = match cached_type {
            FormulaCachedType::Error => Some("#VALUE!".to_owned()),
            _ => None,
        };
        let mut cell = FormulaCell {
            row,
            column,
            formula,
            format_index,
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
            | FormulaCachedType::Empty => {
                self.last_cell = Some(cell.clone());
                Some(cell)
            }
        }
    }

    /// Java `StringRecordHandler` follow-up for a pending string formula.
    pub fn complete_pending_string(&mut self, value: String) -> Option<FormulaCell> {
        let mut cell = self.pending.take()?;
        cell.pending_string = false;
        cell.string_value = Some(value);
        self.last_cell = Some(cell.clone());
        Some(cell)
    }
}

/// BIFF `Formula` record sid. (POI `FormulaRecord.sid`)
pub const FORMULA_SID: u16 = 0x0006;

impl XlsRecordHandler for FormulaRecordHandler {
    /// Java `FormulaRecordHandler.processRecord` — parses coordinates, XF and
    /// every cached-result variant. Formula token text remains a higher layer.
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        if record_sid != FORMULA_SID || data.len() < 14 {
            return;
        }
        let row = u16::from_le_bytes([data[0], data[1]]) as u32;
        let column = u16::from_le_bytes([data[2], data[3]]) as usize;
        let format_index = u16::from_le_bytes([data[4], data[5]]);
        let result = &data[6..14];
        let (cached_type, number_value, bool_value) = if result[6] == 0xFF && result[7] == 0xFF {
            match result[0] {
                0x00 => (FormulaCachedType::String, None, None),
                0x01 => (FormulaCachedType::Boolean, None, Some(result[2] != 0)),
                0x02 => (FormulaCachedType::Error, None, None),
                0x03 => (FormulaCachedType::Empty, None, None),
                _ => (FormulaCachedType::Empty, None, None),
            }
        } else {
            let mut bytes = [0; 8];
            bytes.copy_from_slice(result);
            (
                FormulaCachedType::Numeric,
                Some(f64::from_le_bytes(bytes)),
                None,
            )
        };
        let _ = self.process_formula_with_format(
            row,
            column,
            None,
            cached_type,
            number_value,
            bool_value,
            format_index,
        );
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

    #[test]
    fn process_record_decodes_numeric_boolean_error_and_empty_results() {
        let mut handler = FormulaRecordHandler::new();
        let mut numeric = vec![1, 0, 2, 0, 7, 0];
        numeric.extend_from_slice(&4.5f64.to_le_bytes());
        handler.process_record(FORMULA_SID, &numeric);
        let cell = handler.last_cell.as_ref().expect("numeric formula");
        assert_eq!(cell.cached_type, FormulaCachedType::Numeric);
        assert_eq!(cell.number_value, Some(4.5));
        assert_eq!(cell.format_index, 7);

        for (kind, value, expected) in [
            (0x01, 1, FormulaCachedType::Boolean),
            (0x02, 0x0F, FormulaCachedType::Error),
            (0x03, 0, FormulaCachedType::Empty),
        ] {
            let mut special = vec![1, 0, 2, 0, 0, 0, kind, 0, value, 0, 0, 0, 0xFF, 0xFF];
            handler.process_record(FORMULA_SID, &special);
            assert_eq!(
                handler
                    .last_cell
                    .as_ref()
                    .expect("special formula")
                    .cached_type,
                expected
            );
            special.clear();
        }
    }
}
