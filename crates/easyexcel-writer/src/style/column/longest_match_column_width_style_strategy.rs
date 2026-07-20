//! Mirrors Java `com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy`.

use std::collections::HashMap;
use std::sync::Mutex;

use easyexcel_core::{CellDataType, CellValue, Result, WriteCellContext, WriteHandler};

use crate::style::column::abstract_head_column_width_style_strategy::AbstractHeadColumnWidthStyleStrategy;

/// Maximum Excel column width in character units. (Java `MAX_COLUMN_WIDTH = 255`)
const MAX_COLUMN_WIDTH: u16 = 255;

/// Mirrors Java `LongestMatchColumnWidthStyleStrategy`.
///
/// Java walks rendered cell content after each cell write, measures
/// `String.getBytes().length`, and calls `Sheet.setColumnWidth(col, len * 256)`
/// when a longer value appears. The Rust port:
/// - records UTF-8 byte lengths in [`WriteHandler::after_cell`]
/// - exposes the running max via [`WriteHandler::style_column_width`]
/// - the XLSX write path reapplies those widths after the sheet finishes
///
/// Optional [`Self::with_autofit_fallback`] keeps `worksheet.autofit()` as a
/// secondary path (disabled by default).
pub struct LongestMatchColumnWidthStyleStrategy {
    /// Per-column maximum content length. (Java `cache` / `maxColumnWidthMap`)
    cache: Mutex<HashMap<usize, u16>>,
    /// When true, also request autofit after the sheet write.
    autofit_fallback: bool,
}

impl LongestMatchColumnWidthStyleStrategy {
    /// Creates the strategy with length-based widths only.
    /// (Java `LongestMatchColumnWidthStyleStrategy()`)
    #[must_use]
    pub fn new() -> Self {
        Self {
            cache: Mutex::new(HashMap::new()),
            autofit_fallback: false,
        }
    }

    /// Enables or disables autofit as an optional fallback after length widths.
    #[must_use]
    pub fn with_autofit_fallback(mut self, enabled: bool) -> Self {
        self.autofit_fallback = enabled;
        self
    }

    /// Returns whether autofit fallback is enabled.
    #[must_use]
    pub const fn autofit_fallback(&self) -> bool {
        self.autofit_fallback
    }

    /// Updates the cached max width for one cell. (Java `setColumnWidth` body)
    fn observe_cell(&self, context: &WriteCellContext) {
        let Some(column_width) = data_length(context) else {
            return;
        };
        let column_width = column_width.min(MAX_COLUMN_WIDTH);
        let column_index = usize::from(context.column_index);
        let Ok(mut cache) = self.cache.lock() else {
            return;
        };
        let entry = cache.entry(column_index).or_insert(0);
        if column_width > *entry {
            *entry = column_width;
        }
    }
}

impl Default for LongestMatchColumnWidthStyleStrategy {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for LongestMatchColumnWidthStyleStrategy {
    fn order(&self) -> i32 {
        // Mirror Java `OrderConstant.DEFINE_STYLE` / late column-width apply.
        -50_000
    }

    fn after_cell(&mut self, context: &WriteCellContext) -> Result<()> {
        // Java `AbstractColumnWidthStyleStrategy.afterCellDispose` → `setColumnWidth`
        self.observe_cell(context);
        Ok(())
    }

    fn style_column_width(&self, column_index: usize) -> Option<u16> {
        self.cache
            .lock()
            .ok()
            .and_then(|cache| cache.get(&column_index).copied())
            .filter(|width| *width > 0)
    }

    fn style_auto_column_width(&self) -> bool {
        self.autofit_fallback
    }
}

impl AbstractHeadColumnWidthStyleStrategy for LongestMatchColumnWidthStyleStrategy {
    fn head_column_width(&self, column_index: usize) -> Option<u16> {
        self.style_column_width(column_index)
    }
}

/// Computes the Java-compatible content length for longest-match column width.
///
/// Head cells always use the string/text form. Content cells only measure
/// STRING / BOOLEAN / NUMBER (Java `dataLength` switch); other types return
/// `None` (Java `-1`).
fn data_length(context: &WriteCellContext) -> Option<u16> {
    if context.is_head {
        return byte_len(&context.value.as_text());
    }
    // Java unwraps WriteCellData list; Images/Comment wrap a scalar value.
    let value = match &context.value {
        CellValue::Comment { value, .. } | CellValue::Images { value, .. } => value.as_ref(),
        other => other,
    };
    match value.data_type() {
        CellDataType::String | CellDataType::Boolean | CellDataType::Number => {
            byte_len(&value.as_text())
        }
        _ => None,
    }
}

/// UTF-8 byte length capped to `u16`, approximating Java `String.getBytes().length`.
fn byte_len(text: &str) -> Option<u16> {
    u16::try_from(text.as_bytes().len()).ok()
}
