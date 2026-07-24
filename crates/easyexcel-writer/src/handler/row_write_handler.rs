//! Mirrors Java `com.alibaba.excel.write.handler.RowWriteHandler`.

/// Marks a handler as the Rust counterpart of Java `RowWriteHandler`.
///
/// Implement `before_row_create`, `after_row_create`, and
/// `after_row_dispose` on [`easyexcel_core::WriteHandler`].
pub trait RowWriteHandler: easyexcel_core::WriteHandler {}
