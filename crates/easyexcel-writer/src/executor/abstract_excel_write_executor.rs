//! Mirrors Java `com.alibaba.excel.write.executor.AbstractExcelWriteExecutor`.

use easyexcel_core::WriteContext;

/// Mirrors Java `AbstractExcelWriteExecutor implements ExcelWriteExecutor`.
///
/// The Java side stores a `WriteContext` and exposes the
/// `converterAndSet(CellWriteHandlerContext)` helper. Rust mirrors the
/// field so 1:1 parity holds, but the actual conversion path lives in
/// the derived macro and the writer crate.
pub struct AbstractExcelWriteExecutor<'a> {
    /// Mirrors `AbstractExcelWriteExecutor.writeContext`.
    pub write_context: &'a dyn WriteContext,
}

impl<'a> AbstractExcelWriteExecutor<'a> {
    /// Creates the executor. (Java `AbstractExcelWriteExecutor(WriteContext)`)
    #[must_use]
    pub const fn new(write_context: &'a dyn WriteContext) -> Self {
        Self { write_context }
    }
}
