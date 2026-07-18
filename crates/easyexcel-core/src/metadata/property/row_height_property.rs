//! Mirrors Java `com.alibaba.excel.metadata.property.RowHeightProperty`.

/// Mirrors Java `RowHeightProperty`. (Java `height: Short`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RowHeightProperty {
    /// Row height in points. (Java `getHeight()`)
    pub height: u16,
}

impl RowHeightProperty {
    /// Creates a `RowHeightProperty`. (Java constructor)
    #[must_use]
    pub const fn new(height: u16) -> Self {
        Self { height }
    }
    /// Returns the height. (Java `getHeight()`)
    #[must_use]
    pub const fn height(&self) -> u16 {
        self.height
    }
}
