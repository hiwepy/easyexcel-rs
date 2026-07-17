//! Mirrors Java `com.alibaba.excel.metadata.CellExtra`.

use crate::enum_cell_extra_type::CellExtraType;

/// Extra worksheet information equivalent to Java `EasyExcel`'s `CellExtra`.
///
/// Java carries `rowIndex / columnIndex` plus the interval bounds. Rust keeps
/// the interval bounds as `first_row_index` / `last_row_index` /
/// `first_column_index` / `last_column_index`, while `AnalysisContext` carries
/// the singular cell coordinates, matching how the Java readers forward the
/// event.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellExtra {
    extra_type: CellExtraType,
    text: Option<String>,
    first_row_index: u32,
    last_row_index: u32,
    first_column_index: usize,
    last_column_index: usize,
}

impl CellExtra {
    /// Creates a cell or range extra event. (Java `CellExtra(type, text, firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex)`)
    #[allow(clippy::too_many_arguments)]
    #[must_use]
    pub const fn new(
        extra_type: CellExtraType,
        text: Option<String>,
        first_row_index: u32,
        last_row_index: u32,
        first_column_index: usize,
        last_column_index: usize,
    ) -> Self {
        Self {
            extra_type,
            text,
            first_row_index,
            last_row_index,
            first_column_index,
            last_column_index,
        }
    }

    /// Returns the extra-data kind. (Java `getType()`)
    #[must_use]
    pub const fn extra_type(&self) -> CellExtraType {
        self.extra_type
    }

    /// Returns comment text or hyperlink target; merge events have no text. (Java `getText()`)
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// Returns the first zero-based row index. (Java `getFirstRowIndex()`)
    #[must_use]
    pub const fn first_row_index(&self) -> u32 {
        self.first_row_index
    }

    /// Returns the last zero-based row index. (Java `getLastRowIndex()`)
    #[must_use]
    pub const fn last_row_index(&self) -> u32 {
        self.last_row_index
    }

    /// Returns the first zero-based column index. (Java `getFirstColumnIndex()`)
    #[must_use]
    pub const fn first_column_index(&self) -> usize {
        self.first_column_index
    }

    /// Returns the last zero-based column index. (Java `getLastColumnIndex()`)
    #[must_use]
    pub const fn last_column_index(&self) -> usize {
        self.last_column_index
    }
}
