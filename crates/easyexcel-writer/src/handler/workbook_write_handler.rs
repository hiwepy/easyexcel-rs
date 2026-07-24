//! Mirrors Java `com.alibaba.excel.write.handler.WorkbookWriteHandler`.

/// Marks a handler as the Rust counterpart of Java `WorkbookWriteHandler`.
///
/// The executable `before_workbook_create`, `after_workbook_create`, and
/// `after_workbook_dispose` hooks live on [`easyexcel_core::WriteHandler`].
/// Keeping one object-safe lifecycle trait avoids maintaining a second set of
/// callbacks that the writer cannot discover through `dyn WriteHandler`.
pub trait WorkbookWriteHandler: easyexcel_core::WriteHandler {}
