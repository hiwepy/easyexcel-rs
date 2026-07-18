//! Mirrors Java `com.alibaba.excel.write.handler.impl.DimensionWorkbookWriteHandler`.

use easyexcel_core::WriteWorkbookContext;

/// Mirrors Java `DimensionWorkbookWriteHandler implements WorkbookWriteHandler`.
///
/// Java's handler fixes the `<dimension ref="A1:..."/>` field on
/// `SXSSFWorkbook` because POI's streaming writer skips it. The Rust
/// port delegates this to `rust_xlsxwriter` which always sets the
/// dimension when saving; this marker type exists for parity.
pub struct DimensionWorkbookWriteHandler {
    last_ref: Option<String>,
}

impl DimensionWorkbookWriteHandler {
    /// Creates the handler.
    #[must_use]
    pub const fn new() -> Self {
        Self { last_ref: None }
    }

    /// Returns the last written dimension reference. (Java `getDimension()` step)
    #[must_use]
    pub fn last_ref(&self) -> Option<&str> {
        self.last_ref.as_deref()
    }
}

impl Default for DimensionWorkbookWriteHandler {
    fn default() -> Self {
        Self::new()
    }
}

impl easyexcel_core::WriteHandler for DimensionWorkbookWriteHandler {
    fn after_workbook(&mut self, context: &WriteWorkbookContext) -> easyexcel_core::Result<()> {
        // `rust_xlsxwriter` writes the dimension automatically based on
        // the worksheet bounds. The shim records the path for parity.
        self.last_ref = Some(context.path().display().to_string());
        Ok(())
    }
}