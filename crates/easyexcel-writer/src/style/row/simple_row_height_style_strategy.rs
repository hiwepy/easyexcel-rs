//! Mirrors Java `com.alibaba.excel.write.style.row.SimpleRowHeightStyleStrategy`.

use easyexcel_core::WriteHandler;

use crate::style::row::abstract_row_height_style_strategy::AbstractRowHeightStyleStrategy;

/// Mirrors Java `SimpleRowHeightStyleStrategy`.
pub struct SimpleRowHeightStyleStrategy {
    head_row_height: Option<u16>,
    content_row_height: Option<u16>,
}

impl SimpleRowHeightStyleStrategy {
    /// Creates a strategy with the given row heights. (Java constructor)
    #[must_use]
    pub const fn new(head_row_height: Option<u16>, content_row_height: Option<u16>) -> Self {
        Self {
            head_row_height,
            content_row_height,
        }
    }
}

impl WriteHandler for SimpleRowHeightStyleStrategy {
    fn order(&self) -> i32 {
        -50_000
    }
}

impl AbstractRowHeightStyleStrategy for SimpleRowHeightStyleStrategy {
    fn head_row_height(&self) -> Option<u16> {
        self.head_row_height
    }
    fn content_row_height(&self) -> Option<u16> {
        self.content_row_height
    }
}