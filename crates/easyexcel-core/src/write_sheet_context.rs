//! Mirrors Java `com.alibaba.excel.write.handler.context.SheetWriteHandlerContext`.

/// Worksheet-level write lifecycle context.
///
/// Mirrors Java `SheetWriteHandlerContext` (`writeSheetHolder.getSheetName()`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheetContext {
    sheet_name: String,
}

impl WriteSheetContext {
    /// Creates a worksheet context.
    #[must_use]
    pub fn new(sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
        }
    }

    /// Returns the worksheet name. (Java `WriteSheetHolder.getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
}
