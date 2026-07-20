//! Mirrors Java `com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy`.

use std::collections::HashMap;

use easyexcel_core::WriteHandler;

use crate::style::column::abstract_column_width_style_strategy::AbstractColumnWidthStyleStrategy;

/// Mirrors Java `SimpleColumnWidthStyleStrategy`.
pub struct SimpleColumnWidthStyleStrategy {
    widths: HashMap<usize, u16>,
}

impl SimpleColumnWidthStyleStrategy {
    /// Creates the strategy. (Java `SimpleColumnWidthStyleStrategy()`)
    #[must_use]
    pub fn new() -> Self {
        Self {
            widths: HashMap::new(),
        }
    }

    /// Creates a strategy that applies the same width to every column index
    /// queried later. (Java `SimpleColumnWidthStyleStrategy(Integer columnWidth)`)
    ///
    /// Stores width under key `usize::MAX` as the uniform fallback; callers of
    /// [`AbstractColumnWidthStyleStrategy::column_width`] that pass a concrete
    /// index still win via [`Self::set_column_width`].
    #[must_use]
    pub fn uniform(column_width: u16) -> Self {
        let mut widths = HashMap::new();
        widths.insert(usize::MAX, column_width);
        Self { widths }
    }

    /// Sets a column width. (Java `setColumnWidth(Integer, Integer)`)
    pub fn set_column_width(&mut self, column_index: usize, width: u16) {
        self.widths.insert(column_index, width);
    }
}

impl Default for SimpleColumnWidthStyleStrategy {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteHandler for SimpleColumnWidthStyleStrategy {
    fn order(&self) -> i32 {
        // Java `OrderConstant.DEFINE_STYLE`
        -50_000
    }

    fn style_column_width(&self, column_index: usize) -> Option<u16> {
        AbstractColumnWidthStyleStrategy::column_width(self, column_index)
    }
}

impl AbstractColumnWidthStyleStrategy for SimpleColumnWidthStyleStrategy {
    fn column_width(&self, column_index: usize) -> Option<u16> {
        self.widths
            .get(&column_index)
            .or_else(|| self.widths.get(&usize::MAX))
            .copied()
    }
}
