//! Mirrors Java `com.alibaba.excel.metadata.data.CoordinateData`.

/// Cell coordinates used by Java `CoordinateData` decorations.
///
/// Java uses boxed `Integer`; Rust uses `Option` to express "unspecified" with
/// zero overhead. All four absolute and four relative coordinates follow the
/// Java semantic where zero defers to the relative coordinate.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[allow(clippy::struct_field_names)]
pub struct CoordinateData {
    first_row_index: Option<u32>,
    first_column_index: Option<u16>,
    last_row_index: Option<u32>,
    last_column_index: Option<u16>,
    relative_first_row_index: Option<i32>,
    relative_first_column_index: Option<i32>,
    relative_last_row_index: Option<i32>,
    relative_last_column_index: Option<i32>,
}

impl CoordinateData {
    /// Creates coordinates that default to the decorated cell.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            first_row_index: None,
            first_column_index: None,
            last_row_index: None,
            last_column_index: None,
            relative_first_row_index: None,
            relative_first_column_index: None,
            relative_last_row_index: None,
            relative_last_column_index: None,
        }
    }

    /// Sets the absolute first row. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn first_row_index(mut self, value: u32) -> Self {
        self.first_row_index = Some(value);
        self
    }

    /// Sets the absolute first column. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn first_column_index(mut self, value: u16) -> Self {
        self.first_column_index = Some(value);
        self
    }

    /// Sets the absolute last row. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn last_row_index(mut self, value: u32) -> Self {
        self.last_row_index = Some(value);
        self
    }

    /// Sets the absolute last column. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn last_column_index(mut self, value: u16) -> Self {
        self.last_column_index = Some(value);
        self
    }

    /// Sets the first row relative to the decorated cell.
    #[must_use]
    pub const fn relative_first_row_index(mut self, value: i32) -> Self {
        self.relative_first_row_index = Some(value);
        self
    }

    /// Sets the first column relative to the decorated cell.
    #[must_use]
    pub const fn relative_first_column_index(mut self, value: i32) -> Self {
        self.relative_first_column_index = Some(value);
        self
    }

    /// Sets the last row relative to the decorated cell.
    #[must_use]
    pub const fn relative_last_row_index(mut self, value: i32) -> Self {
        self.relative_last_row_index = Some(value);
        self
    }

    /// Sets the last column relative to the decorated cell.
    #[must_use]
    pub const fn relative_last_column_index(mut self, value: i32) -> Self {
        self.relative_last_column_index = Some(value);
        self
    }

    /// Returns the absolute first row. (Java `getFirstRowIndex()`)
    #[must_use]
    pub const fn get_first_row_index(self) -> Option<u32> {
        self.first_row_index
    }

    /// Returns the absolute first column. (Java `getFirstColumnIndex()`)
    #[must_use]
    pub const fn get_first_column_index(self) -> Option<u16> {
        self.first_column_index
    }

    /// Returns the absolute last row. (Java `getLastRowIndex()`)
    #[must_use]
    pub const fn get_last_row_index(self) -> Option<u32> {
        self.last_row_index
    }

    /// Returns the absolute last column. (Java `getLastColumnIndex()`)
    #[must_use]
    pub const fn get_last_column_index(self) -> Option<u16> {
        self.last_column_index
    }

    /// Returns the relative first row. (Java `getRelativeFirstRowIndex()`)
    #[must_use]
    pub const fn get_relative_first_row_index(self) -> Option<i32> {
        self.relative_first_row_index
    }

    /// Returns the relative first column. (Java `getRelativeFirstColumnIndex()`)
    #[must_use]
    pub const fn get_relative_first_column_index(self) -> Option<i32> {
        self.relative_first_column_index
    }

    /// Returns the relative last row. (Java `getRelativeLastRowIndex()`)
    #[must_use]
    pub const fn get_relative_last_row_index(self) -> Option<i32> {
        self.relative_last_row_index
    }

    /// Returns the relative last column. (Java `getRelativeLastColumnIndex()`)
    #[must_use]
    pub const fn get_relative_last_column_index(self) -> Option<i32> {
        self.relative_last_column_index
    }
}
