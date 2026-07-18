//! Mirrors Java `com.alibaba.excel.write.handler.context.CellWriteHandlerContext`.

use easyexcel_core::{CellValue, WriteCellContext};

/// Mirrors Java `CellWriteHandlerContext` (132-line POI-backed value object).
///
/// The Java side carries a POI `Cell`, `Row`, holder references, and
/// the resolved `WriteCellData`. Rust collapses the POI references
/// and keeps a flat context struct so handlers receive the
/// same data shape.
#[derive(Debug, Clone)]
pub struct CellWriteHandlerContext {
    /// Mirrors the resolved cell value. (Java `getFirstCellData()`)
    pub first_cell_data: Option<CellValue>,
    /// The cell context the handler is operating on.
    pub cell: WriteCellContext,
}

impl CellWriteHandlerContext {
    /// Returns the first cell data, if any. (Java `getFirstCellData()`)
    #[must_use]
    pub const fn first_cell_data(&self) -> Option<&CellValue> {
        self.first_cell_data.as_ref()
    }

    /// Returns the inner cell context. (Java `getCell()` step)
    #[must_use]
    pub const fn cell(&self) -> &WriteCellContext {
        &self.cell
    }
}