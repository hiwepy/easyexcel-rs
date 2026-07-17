//! Mirrors Java `com.alibaba.excel.converters.ReadConverterContext`.

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_column::ExcelColumn;
use crate::formula_data::FormulaData;

/// Context supplied to a custom cell-to-Rust converter.
///
/// Mirrors Java `ReadConverterContext<T>(readCellData, contentProperty,
/// analysisContext)`. Rust drops the `ReadCellData` wrapper and stores
/// `&CellValue` directly to avoid cloning the entire cell envelope.
#[derive(Debug, Clone, Copy)]
pub struct ReadConverterContext<'a> {
    cell: Option<&'a CellValue>,
    formula: Option<&'a FormulaData>,
    column: &'a ExcelColumn,
    context: &'a ConvertContext,
}

impl<'a> ReadConverterContext<'a> {
    /// Creates a read conversion context. (Java `@AllArgsConstructor`)
    #[must_use]
    pub const fn new(
        cell: Option<&'a CellValue>,
        column: &'a ExcelColumn,
        context: &'a ConvertContext,
    ) -> Self {
        Self {
            cell,
            formula: None,
            column,
            context,
        }
    }

    /// Creates a read conversion context with optional formula metadata.
    /// Mirrors Java's ability to expose `formulaData` from `ReadCellData`.
    #[must_use]
    pub const fn with_formula(
        cell: Option<&'a CellValue>,
        formula: Option<&'a FormulaData>,
        column: &'a ExcelColumn,
        context: &'a ConvertContext,
    ) -> Self {
        Self {
            cell,
            formula,
            column,
            context,
        }
    }

    /// Returns the source cell, or `None` when it is physically absent. (Java `getReadCellData()`)
    #[must_use]
    pub const fn cell(&self) -> Option<&'a CellValue> {
        self.cell
    }

    /// Returns formula metadata when the source cell contains a formula. (Java `ReadCellData.getFormulaData()`)
    #[must_use]
    pub const fn formula(&self) -> Option<&'a FormulaData> {
        self.formula
    }

    /// Returns the field's static column metadata. (Java `getContentProperty()`)
    #[must_use]
    pub const fn column(&self) -> &'a ExcelColumn {
        self.column
    }

    /// Returns the resolved row, column, field, and format information. (Java `getAnalysisContext()`)
    #[must_use]
    pub const fn convert_context(&self) -> &'a ConvertContext {
        self.context
    }
}
