//! Mirrors Java `com.alibaba.excel.write.handler.impl.DefaultRowWriteHandler`.

use crate::WriteHandler;
use easyexcel_core::WriteContext;
use easyexcel_core::WriteSheetContext;

/// Mirrors Java `DefaultRowWriteHandler extends AbstractRowWriteHandler`.
///
/// Java's handler simply hooks `beforeSheetCreate` to freeze the first
/// row, which `rust_xlsxwriter` does automatically via
/// `worksheet.set_freeze_panes(...)`. The Rust shim is preserved so
/// 1:1 code references resolve.
pub struct DefaultRowWriteHandler {
    frozen: bool,
}

impl DefaultRowWriteHandler {
    /// Creates the handler.
    #[must_use]
    pub const fn new() -> Self {
        Self { frozen: false }
    }

    /// Returns whether the first row is frozen. (Java `getFreeze()` equivalent)
    #[must_use]
    pub const fn frozen(&self) -> bool {
        self.frozen
    }
}

impl Default for DefaultRowWriteHandler {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for DefaultRowWriteHandler {
    fn before_sheet(&mut self, _context: &WriteSheetContext) -> easyexcel_core::Result<()> {
        // The actual freeze is performed in [`crate::ExcelWriter::write`]
        // by inspecting `WriteOptions.freeze_head`.
        self.frozen = true;
        Ok(())
    }
}

/// Mirrors the Java constructor pattern that received a
/// `WriteContext` for back-reference. Kept for parity.
pub fn new_default_row_write_handler(_ctx: &dyn WriteContext) -> DefaultRowWriteHandler {
    DefaultRowWriteHandler::new()
}