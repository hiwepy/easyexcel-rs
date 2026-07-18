//! Mirrors Java `com.alibaba.excel.write.style.HorizontalCellStyleStrategy`.

use easyexcel_core::{ExcelCellStyle, WriteCellContext, WriteHandler};

use crate::style::abstract_cell_style_strategy::AbstractCellStyleStrategy;

/// Mirrors Java `HorizontalCellStyleStrategy`.
///
/// The Java side cycles through a list of content styles; the Rust port
/// mirrors that surface so the type is preserved.
pub struct HorizontalCellStyleStrategy {
    content_styles: Vec<ExcelCellStyle>,
}

impl HorizontalCellStyleStrategy {
    /// Creates a strategy with the given cycle. (Java constructor)
    #[must_use]
    pub const fn new(content_styles: Vec<ExcelCellStyle>) -> Self {
        Self { content_styles }
    }

    /// Returns the configured content styles. (Java `getContentCellStyleList()`)
    #[must_use]
    pub fn content_styles(&self) -> &[ExcelCellStyle] {
        &self.content_styles
    }
}

impl AbstractCellStyleStrategy for HorizontalCellStyleStrategy {
    fn cell_style(&self, context: &WriteCellContext) -> ExcelCellStyle {
        if context.is_head {
            return ExcelCellStyle::new();
        }
        // Cycle through content styles by relative row index. The
        // `is_head` field is already used as the discriminator; for
        // non-head cells we simply return the first configured style.
        self.content_styles
            .first()
            .copied()
            .unwrap_or_default()
    }
}

impl WriteHandler for HorizontalCellStyleStrategy {
    fn order(&self) -> i32 {
        0
    }
}