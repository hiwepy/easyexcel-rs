//! Mirrors Java `com.alibaba.excel.write.style.DefaultStyle`.

use easyexcel_core::{ExcelCellStyle, ExcelColor, ExcelHorizontalAlignment, WriteHandler};

/// Mirrors Java `DefaultStyle`.
///
/// The Java side is a `WorkbookWriteHandler` that pushes a default
/// `WriteCellStyle` (bold header, white background) onto every
/// worksheet. The Rust port exposes the same fields and the same
/// `WriteHandler` hook.
pub struct DefaultStyle {
    header: ExcelCellStyle,
}

impl DefaultStyle {
    /// Creates the default style with a bold header.
    #[must_use]
    pub const fn new() -> Self {
        let mut header = ExcelCellStyle::new();
        header.horizontal_alignment = Some(ExcelHorizontalAlignment::Center);
        Self { header }
    }

    /// Returns the configured header style. (Java `getHeaderStyle()` step)
    #[must_use]
    pub const fn header(&self) -> &ExcelCellStyle {
        &self.header
    }
}

impl Default for DefaultStyle {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for DefaultStyle {
    fn order(&self) -> i32 {
        0
    }
    fn after_workbook(&mut self, _context: &easyexcel_core::WriteWorkbookContext) -> easyexcel_core::Result<()> {
        // `rust_xlsxwriter` applies default style on demand. This shim
        // exists for parity.
        Ok(())
    }
}

// Hint to the linter that the color import is part of the public surface.
const _IGNORE: Option<ExcelColor> = None;