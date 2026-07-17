//! Mirrors Java `com.alibaba.excel.write.handler.context.RowWriteHandlerContext`.

/// Row-level write lifecycle context.
///
/// Mirrors Java `RowWriteHandlerContext` (`writeSheetHolder`, `writeTableHolder`,
/// `rowIndex`, `relativeRowIndex`, `head`). Rust keeps only the fields a
/// handler needs and drops the `Row` POI object because `rust_xlsxwriter`
/// does not expose it for handler interception.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteRowContext {
    /// Worksheet name.
    pub sheet_name: String,
    /// Physical zero-based row index.
    pub row_index: u32,
    /// Whether this is a header row.
    pub is_head: bool,
}
