//! Cell value editor trait — mirrors hutool `cn.hutool.poi.excel.cell.CellEditor`.
//!
//! In hutool, `CellEditor` transforms cell values during reading.
//! In easyexcel-rust, this can be registered on the reader builder and
//! applied before the value reaches the `ReadListener`.

use crate::CellValue;

/// Transforms a cell value during reading.
///
/// Mirrors hutool `CellEditor` interface:
/// ```java
/// public interface CellEditor {
///     Object edit(Cell cell, Object value);
/// }
/// ```
///
/// In Rust, the `Cell` object is replaced by the `CellValue` since
/// we don't have a POI cell handle.
pub trait CellEditor: Send + Sync {
    /// Transforms a cell value before it reaches the listener.
    fn edit(&self, original: &CellValue, sheet_name: &str, row: u32, col: u32) -> CellValue;
}

/// Trims whitespace from string cell values.
///
/// Mirrors hutool `TrimEditor`.
/// Note: easyexcel-rust has `auto_trim(true)` which does this globally
/// without needing a CellEditor. This editor is for selective trimming.
#[derive(Debug, Default, Clone)]
pub struct TrimEditor;

impl CellEditor for TrimEditor {
    fn edit(&self, original: &CellValue, _sheet_name: &str, _row: u32, _col: u32) -> CellValue {
        match original {
            CellValue::String(s) => CellValue::String(s.trim().to_owned()),
            other => other.clone(),
        }
    }
}

/// Converts numeric (Int/Float/Decimal) cell values to integers by truncation.
///
/// Mirrors hutool `NumericToIntEditor`.
#[derive(Debug, Default, Clone)]
pub struct NumericToIntEditor;

impl CellEditor for NumericToIntEditor {
    fn edit(&self, original: &CellValue, _sheet_name: &str, _row: u32, _col: u32) -> CellValue {
        match original {
            CellValue::Int(n) => CellValue::Int(*n),
            CellValue::Float(f) => CellValue::Int(*f as i64),
            CellValue::Decimal(d) => {
                let s = d.to_string();
                if let Ok(n) = s.parse::<i64>() {
                    CellValue::Int(n)
                } else {
                    CellValue::Int(0)
                }
            }
            CellValue::Bool(b) => CellValue::Int(if *b { 1 } else { 0 }),
            CellValue::String(s) => {
                if let Ok(n) = s.trim().parse::<i64>() {
                    CellValue::Int(n)
                } else {
                    original.clone()
                }
            }
            other => other.clone(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn trim_editor_strips_whitespace() {
        let editor = TrimEditor;
        let result = editor.edit(&CellValue::String("  hello  ".into()), "", 0, 0);
        assert_eq!(result, CellValue::String("hello".into()));
    }

    #[test]
    fn trim_editor_preserves_non_string() {
        let editor = TrimEditor;
        let result = editor.edit(&CellValue::Int(42), "", 0, 0);
        assert_eq!(result, CellValue::Int(42));
    }

    #[test]
    fn numeric_editor_converts_float_to_int() {
        let editor = NumericToIntEditor;
        assert_eq!(editor.edit(&CellValue::Float(3.14), "", 0, 0), CellValue::Int(3));
        assert_eq!(editor.edit(&CellValue::Int(42), "", 0, 0), CellValue::Int(42));
        assert_eq!(editor.edit(&CellValue::Bool(true), "", 0, 0), CellValue::Int(1));
        assert_eq!(editor.edit(&CellValue::String("99".into()), "", 0, 0), CellValue::Int(99));
    }
}
