//! Mirrors Java `com.alibaba.excel.write.handler.context.CellWriteHandlerContext`.

use crate::cell_value::CellValue;

/// Mutable cell-level write lifecycle context.
///
/// Mirrors Java `CellWriteHandlerContext` (13 fields). Rust keeps only the
/// fields a handler actually mutates and exposes `skip: bool` so handlers
/// can suppress writing a cell without juggling the underlying POI types.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellContext {
    /// Worksheet name.
    pub sheet_name: String,
    /// Physical zero-based row index.
    pub row_index: u32,
    /// Physical zero-based column index.
    pub column_index: u16,
    /// Rust field name, when backed by a typed column.
    pub field: Option<&'static str>,
    /// Whether this is a header cell.
    pub is_head: bool,
    /// Value that will be written. A handler may replace it.
    pub value: CellValue,
    /// A handler may set this to suppress the physical cell.
    pub skip: bool,
}
