//! Mirrors Java `com.alibaba.excel.write.handler.CellWriteHandler`.

/// Marks a handler as the Rust counterpart of Java `CellWriteHandler`.
///
/// Implement `before_cell_create`, `after_cell_create`,
/// `after_cell_data_converted`, and `after_cell_dispose` on
/// [`easyexcel_core::WriteHandler`].
pub trait CellWriteHandler: easyexcel_core::WriteHandler {}
