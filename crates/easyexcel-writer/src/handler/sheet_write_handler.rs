//! Mirrors Java `com.alibaba.excel.write.handler.SheetWriteHandler`.

/// Marks a handler as the Rust counterpart of Java `SheetWriteHandler`.
///
/// Implement `before_sheet_create` and `after_sheet_create` on
/// [`easyexcel_core::WriteHandler`]; those are the hooks invoked by the
/// writer engine.
pub trait SheetWriteHandler: easyexcel_core::WriteHandler {}
