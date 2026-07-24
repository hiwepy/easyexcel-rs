//! Mirrors Java `com.alibaba.excel.write.handler.impl.DefaultWriteSheetHandler`.

use easyexcel_core::{Result, WriteSheetContext};

use crate::WriteHandler;

/// Mirrors Java `DefaultWriteSheetHandler`.
///
/// Tracks whether the sheet has been initialized for writing so the
/// builder can defer dimension calculation until the first row arrives.
pub struct DefaultWriteSheetHandler {
    initialized: bool,
}

impl DefaultWriteSheetHandler {
    /// Creates the handler. (Java `DefaultWriteSheetHandler()`)
    #[must_use]
    pub const fn new() -> Self {
        Self { initialized: false }
    }

    /// Returns whether the sheet has been initialized.
    #[must_use]
    pub const fn initialized(&self) -> bool {
        self.initialized
    }
}

impl Default for DefaultWriteSheetHandler {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for DefaultWriteSheetHandler {
    fn after_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        // Java: `DefaultWriteSheetHandler.afterSheetCreate` just marks
        // the sheet as initialized so subsequent rows can be appended.
        self.initialized = true;
        Ok(())
    }
}
