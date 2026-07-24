//! Mirrors Java `com.alibaba.excel.write.handler.impl.DefaultWriteWorkbookHandler`.

use easyexcel_core::{Result, WriteWorkbookContext};

use crate::WriteHandler;

/// Mirrors Java `DefaultWriteWorkbookHandler`.
///
/// Tracks whether the workbook has been initialized for writing so the
/// builder can defer dimension calculation until the first sheet
/// arrives.
pub struct DefaultWriteWorkbookHandler {
    initialized: bool,
}

impl DefaultWriteWorkbookHandler {
    /// Creates the handler. (Java `DefaultWriteWorkbookHandler()`)
    #[must_use]
    pub const fn new() -> Self {
        Self { initialized: false }
    }

    /// Returns whether the workbook has been initialized.
    #[must_use]
    pub const fn initialized(&self) -> bool {
        self.initialized
    }
}

impl Default for DefaultWriteWorkbookHandler {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for DefaultWriteWorkbookHandler {
    fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        // Java: `DefaultWriteWorkbookHandler.beforeWorkbookCreate` just
        // marks the workbook as initialized so subsequent sheet
        // creation can proceed.
        self.initialized = true;
        Ok(())
    }
}
