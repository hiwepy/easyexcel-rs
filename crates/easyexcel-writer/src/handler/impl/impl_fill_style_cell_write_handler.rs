//! Mirrors Java `com.alibaba.excel.write.handler.impl.FillStyleCellWriteHandler`.

use easyexcel_core::{ExcelCellStyle, WriteCellContext, WriteHandler};

/// Mirrors Java `FillStyleCellWriteHandler`.
///
/// Java's handler keeps the originating template cell style when filling
/// a template. The Rust port delegates this to the template fill code
/// in `easyexcel-template`; this type exists for parity.
pub struct FillStyleCellWriteHandler {
    last_style: Option<ExcelCellStyle>,
    ignore: bool,
}

impl FillStyleCellWriteHandler {
    /// Creates the handler.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            last_style: None,
            ignore: false,
        }
    }

    /// Returns the last observed style. (Java `getOriginCellStyle()` step)
    #[must_use]
    pub const fn last_style(&self) -> Option<&ExcelCellStyle> {
        self.last_style.as_ref()
    }

    /// Sets the `ignoreFillStyle` flag. (Java `setIgnoreFillStyle`)
    pub fn set_ignore(&mut self, ignore: bool) {
        self.ignore = ignore;
    }
}

impl Default for FillStyleCellWriteHandler {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for FillStyleCellWriteHandler {
    fn before_cell(&mut self, _context: &mut WriteCellContext) -> easyexcel_core::Result<()> {
        // `easyexcel-template` consults `ignoreFillStyle` when emitting
        // each cell. The handler is preserved here for parity.
        Ok(())
    }
}
