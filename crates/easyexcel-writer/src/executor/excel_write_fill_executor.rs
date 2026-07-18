//! Mirrors Java `com.alibaba.excel.write.executor.ExcelWriteFillExecutor`.

use easyexcel_core::WriteContext;

use crate::executor::abstract_excel_write_executor::AbstractExcelWriteExecutor;

/// Mirrors Java `ExcelWriteFillExecutor extends AbstractExcelWriteExecutor`.
///
/// The Java side handles `{key}` / `{prefix.field}` placeholder
/// substitution. The Rust port implements the same algorithm in
/// `easyexcel_template::ExcelTemplateWriter`; this struct exists for
/// 1:1 API parity.
pub struct ExcelWriteFillExecutor<'a> {
    inner: AbstractExcelWriteExecutor<'a>,
}

impl<'a> ExcelWriteFillExecutor<'a> {
    /// Creates the executor. (Java `ExcelWriteFillExecutor(WriteContext)`)
    #[must_use]
    pub const fn new(write_context: &'a dyn WriteContext) -> Self {
        Self {
            inner: AbstractExcelWriteExecutor::new(write_context),
        }
    }

    /// Returns the inner `WriteContext`. (Java `getWriteContext()` step)
    #[must_use]
    pub const fn write_context(&self) -> &dyn WriteContext {
        self.inner.write_context
    }
}