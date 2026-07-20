//! Mirrors Java `com.alibaba.excel.write.metadata.fill.AnalysisCell`.

use crate::WriteTemplateAnalysisCellType;

/// Template placeholder discovered while filling data.
///
/// Rust port of Java `AnalysisCell`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnalysisCell {
    /// Zero-based column index. (Java `columnIndex`)
    pub column_index: i32,
    /// Zero-based row index. (Java `rowIndex`)
    pub row_index: i32,
    /// Placeholder variables such as `{name}`. (Java `variableList`)
    pub variable_list: Vec<String>,
    /// Prepared data tokens. (Java `prepareDataList`)
    pub prepare_data_list: Vec<String>,
    /// Whether the cell contains exactly one variable. (Java `onlyOneVariable`)
    pub only_one_variable: Option<bool>,
    /// Template cell kind. (Java `cellType`)
    pub cell_type: WriteTemplateAnalysisCellType,
    /// Prefix before the first variable. (Java `prefix`)
    pub prefix: Option<String>,
    /// Whether this is the first row of a collection block. (Java `firstRow`)
    pub first_row: Option<bool>,
}

impl AnalysisCell {
    /// Creates a common template cell. (Java `initAnalysisCell`)
    #[must_use]
    pub fn new(row_index: i32, column_index: i32) -> Self {
        Self {
            column_index,
            row_index,
            variable_list: Vec::new(),
            prepare_data_list: Vec::new(),
            only_one_variable: None,
            cell_type: WriteTemplateAnalysisCellType::Common,
            prefix: None,
            first_row: None,
        }
    }

    /// Returns the column index. (Java `getColumnIndex()`)
    #[must_use]
    pub const fn column_index(&self) -> i32 {
        self.column_index
    }

    /// Returns the row index. (Java `getRowIndex()`)
    #[must_use]
    pub const fn row_index(&self) -> i32 {
        self.row_index
    }

    /// Returns the template cell kind. (Java `getCellType()`)
    #[must_use]
    pub const fn cell_type(&self) -> WriteTemplateAnalysisCellType {
        self.cell_type
    }
}
