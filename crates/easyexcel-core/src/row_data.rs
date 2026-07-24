//! Mirrors the union of `ReadRowHolder.cellMap` / `CellExtra` /
//! `currentRowAnalysisResult` aggregated into one row payload.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use bigdecimal::BigDecimal;

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::dynamic_value::DynamicValue;
use crate::enum_read_default_return::ReadDefaultReturn;
use crate::excel_column::ExcelColumn;
use crate::formula_data::FormulaData;

/// A physical row plus resolved header positions.
///
/// Java distributes these across `ReadRowHolder` (current row), `CellData`
/// (per-cell scalars), and `AnalysisContext` (current sheet / row index).
/// Rust fuses them into a single value that travels from `XlsxSaxAnalyser`
/// through `T::from_row_with_converters` and into listener callbacks.
#[derive(Debug, Clone)]
pub struct RowData {
    sheet_name: String,
    row_index: u32,
    cells: Vec<CellValue>,
    headers: Arc<HashMap<String, usize>>,
    formulas: HashMap<usize, FormulaData>,
    display_values: HashMap<usize, String>,
    decimal_values: HashMap<usize, BigDecimal>,
    present_columns: HashSet<usize>,
    read_default_return: ReadDefaultReturn,
    use_1904_windowing: bool,
}

impl RowData {
    /// Creates row data. (Java `ReadRowHolder(rowIndex, rowType, globalConfiguration, cellMap)` subset)
    #[must_use]
    pub fn new(
        sheet_name: impl Into<String>,
        row_index: u32,
        cells: Vec<CellValue>,
        headers: Arc<HashMap<String, usize>>,
    ) -> Self {
        let present_columns = (0..cells.len()).collect();
        Self {
            sheet_name: sheet_name.into(),
            row_index,
            cells,
            headers,
            formulas: HashMap::new(),
            display_values: HashMap::new(),
            decimal_values: HashMap::new(),
            present_columns,
            read_default_return: ReadDefaultReturn::default(),
            use_1904_windowing: false,
        }
    }

    /// Attaches formula metadata indexed by zero-based physical column. (Java `CellData.formulaData`)
    #[must_use]
    pub fn with_formulas(mut self, formulas: HashMap<usize, FormulaData>) -> Self {
        self.formulas = formulas;
        self
    }

    /// Attaches Java-compatible formatted display text by physical column index. (Java `CellData.stringValue`)
    #[must_use]
    pub fn with_display_values(mut self, display_values: HashMap<usize, String>) -> Self {
        self.display_values = display_values;
        self
    }

    /// Attaches exact OOXML decimal values by physical column index. (Java `CellData.numberValue`)
    #[must_use]
    pub fn with_decimal_values(mut self, decimal_values: HashMap<usize, BigDecimal>) -> Self {
        self.decimal_values = decimal_values;
        self
    }

    /// Attaches the physical columns that were explicitly present in the source.
    #[must_use]
    pub fn with_present_columns(mut self, present_columns: HashSet<usize>) -> Self {
        self.present_columns = present_columns;
        self
    }

    /// Selects the Java-compatible no-model return mode. (Java `ReadDefaultReturnEnum`)
    #[must_use]
    pub const fn with_read_default_return(mut self, mode: ReadDefaultReturn) -> Self {
        self.read_default_return = mode;
        self
    }

    /// Selects Excel's 1904 numeric date system for field conversion.
    #[must_use]
    pub const fn with_use_1904_windowing(mut self, enabled: bool) -> Self {
        self.use_1904_windowing = enabled;
        self
    }

    /// Returns the physical zero-based row index. (Java `ReadRowHolder.getRowIndex()`)
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the sheet name. (Java `ReadRowHolder.sheetName`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Resolves a cell using Java `EasyExcel`'s index-before-name priority.
    ///
    /// Mirrors Java `AnalysisContext.readRowHolder().getCell(column)` semantics
    /// in the `ModelBuildEventListener` (`buildUserModel`).
    #[must_use]
    pub fn cell(&self, column: &ExcelColumn) -> Option<&CellValue> {
        let index = column
            .index
            .or_else(|| self.headers.get(column.name).copied())?;
        self.cells.get(index)
    }

    /// Resolves formula metadata using the same index-before-name priority as [`Self::cell`].
    #[must_use]
    pub fn formula(&self, column: &ExcelColumn) -> Option<&FormulaData> {
        let index = column
            .index
            .or_else(|| self.headers.get(column.name).copied())?;
        self.formulas.get(&index)
    }

    /// Returns POI-compatible display text retained for a numeric source cell.
    #[must_use]
    pub fn display_value(&self, column: &ExcelColumn) -> Option<&str> {
        let index = column
            .index
            .or_else(|| self.headers.get(column.name).copied())?;
        self.display_values.get(&index).map(String::as_str)
    }

    /// Returns the exact decimal token retained from OOXML for a numeric cell.
    #[must_use]
    pub fn decimal_value(&self, column: &ExcelColumn) -> Option<&BigDecimal> {
        let index = column
            .index
            .or_else(|| self.headers.get(column.name).copied())?;
        self.decimal_values.get(&index)
    }

    /// Dynamic-row support: maximum physical column touched by either headers or cells.
    pub(crate) fn dynamic_width(&self) -> usize {
        let head_width = self
            .headers
            .values()
            .copied()
            .max()
            .map_or(0, |index| index.saturating_add(1));
        self.cells.len().max(head_width)
    }

    /// Dynamic-row support: produce a `DynamicValue` for a column.
    pub(crate) fn dynamic_cell(&self, column_index: usize) -> DynamicValue {
        if !self.present_columns.contains(&column_index) {
            return DynamicValue::Null;
        }
        let raw_value = self
            .cells
            .get(column_index)
            .cloned()
            .unwrap_or(CellValue::Empty);
        let raw_value = if matches!(raw_value, CellValue::Int(_) | CellValue::Float(_)) {
            self.decimal_values
                .get(&column_index)
                .cloned()
                .map_or(raw_value, CellValue::Decimal)
        } else {
            raw_value
        };
        let data = actual_cell_value(&raw_value);
        let display_value = self
            .display_values
            .get(&column_index)
            .cloned()
            .unwrap_or_else(|| raw_value.as_text());
        match self.read_default_return {
            ReadDefaultReturn::String => DynamicValue::String(display_value),
            ReadDefaultReturn::ActualData => DynamicValue::ActualData(data),
            ReadDefaultReturn::ReadCellData => {
                DynamicValue::ReadCellData(crate::read_cell_data::ReadCellData::new(
                    self.row_index,
                    column_index,
                    raw_value,
                    data,
                    display_value,
                    self.formulas.get(&column_index).cloned(),
                ))
            }
        }
    }

    /// Creates a conversion context for a column. (Java `ReadConverterContext` constructor)
    #[must_use]
    pub fn convert_context(&self, column: &ExcelColumn) -> ConvertContext {
        let column_index = column
            .index
            .or_else(|| self.headers.get(column.name).copied());
        ConvertContext {
            sheet_name: self.sheet_name.clone(),
            row_index: self.row_index,
            column_index,
            field: column.field,
            format: column.format,
            use_1904_windowing: column.use_1904_windowing.unwrap_or(self.use_1904_windowing),
        }
    }
}

pub(crate) fn actual_cell_value(value: &CellValue) -> CellValue {
    match value {
        CellValue::Empty => CellValue::String(String::new()),
        CellValue::Error(value) => CellValue::String(value.clone()),
        value => value.clone(),
    }
}
