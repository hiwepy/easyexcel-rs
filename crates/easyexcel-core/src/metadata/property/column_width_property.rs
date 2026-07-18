//! Mirrors Java `com.alibaba.excel.metadata.property.ColumnWidthProperty`.

/// Mirrors Java `ColumnWidthProperty`. (Java `width: Integer`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ColumnWidthProperty {
    /// Column width in Excel character units. (Java `getWidth()`)
    pub width: u16,
}

impl ColumnWidthProperty {
    /// Creates a `ColumnWidthProperty`. (Java constructor)
    #[must_use]
    pub const fn new(width: u16) -> Self {
        Self { width }
    }
    /// Returns the width. (Java `getWidth()`)
    #[must_use]
    pub const fn width(&self) -> u16 {
        self.width
    }
}
