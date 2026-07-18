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
        -50_000
    }
}

impl AbstractColumnWidthStyleStrategy for SimpleColumnWidthStyleStrategy {
    fn column_width(&self, column_index: usize) -> Option<u16> {
        self.widths.get(&column_index).copied()
    }
}