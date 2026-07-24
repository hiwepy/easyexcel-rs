//! Mirrors Java `com.alibaba.excel.converters.ReadConverterContext`.

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_column::ExcelColumn;
use crate::formula_data::FormulaData;
use bigdecimal::BigDecimal;

/// Context supplied to a custom cell-to-Rust converter.
///
/// Mirrors Java `ReadConverterContext<T>(readCellData, contentProperty,
/// analysisContext)`. Rust drops the `ReadCellData` wrapper and stores
/// `&CellValue` directly to avoid cloning the entire cell envelope.
#[derive(Debug, Clone, Copy)]
pub struct ReadConverterContext<'a> {
    cell: Option<&'a CellValue>,
    formula: Option<&'a FormulaData>,
    display_value: Option<&'a str>,
    decimal_value: Option<&'a BigDecimal>,
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
            display_value: None,
            decimal_value: None,
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
            display_value: None,
            decimal_value: None,
            column,
            context,
        }
    }

    /// Creates a context with the full scalar metadata retained by Java `ReadCellData`.
    ///
    /// `display_value` mirrors `ReadCellData.stringValue` after POI
    /// `DataFormatter`; `decimal_value` mirrors the exact
    /// `ReadCellData.numberValue` parsed from OOXML rather than its `f64`
    /// transport representation.
    #[must_use]
    pub const fn with_cell_metadata(
        cell: Option<&'a CellValue>,
        formula: Option<&'a FormulaData>,
        display_value: Option<&'a str>,
        decimal_value: Option<&'a BigDecimal>,
        column: &'a ExcelColumn,
        context: &'a ConvertContext,
    ) -> Self {
        Self {
            cell,
            formula,
            display_value,
            decimal_value,
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

    /// Returns the Excel/POI-compatible rendered text when the reader retained it.
    #[must_use]
    pub const fn display_value(&self) -> Option<&'a str> {
        self.display_value
    }

    /// Returns the exact decimal token retained from the source workbook.
    #[must_use]
    pub const fn decimal_value(&self) -> Option<&'a BigDecimal> {
        self.decimal_value
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
