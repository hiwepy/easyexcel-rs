//! Mirrors Java `com.alibaba.excel.write.executor.ExcelWriteAddExecutor`.

use easyexcel_core::WriteContext;

use crate::executor::abstract_excel_write_executor::AbstractExcelWriteExecutor;

/// Mirrors Java `ExcelWriteAddExecutor extends AbstractExcelWriteExecutor`.
///
/// The Java side holds `add(Collection<?>)`, `addOneRowOfDataToExcel`,
/// `addBasicTypeToExcel`, and `addJavaObjectToExcel`. Rust delegates the
/// heavy lifting to the `rust_xlsxwriter` bindings and the
/// `#[derive(ExcelRow)]` proc-macro; this struct exists so the 1:1 API
/// class name is preserved.
pub struct ExcelWriteAddExecutor<'a> {
    inner: AbstractExcelWriteExecutor<'a>,
}

impl<'a> ExcelWriteAddExecutor<'a> {
    /// Creates the executor. (Java `ExcelWriteAddExecutor(WriteContext)`)
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