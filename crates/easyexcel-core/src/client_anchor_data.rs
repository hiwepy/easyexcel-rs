//! Mirrors Java `com.alibaba.excel.metadata.data.ClientAnchorData`.

use crate::anchor_type::AnchorType;
use crate::coordinate_data::CoordinateData;

/// Client-anchor margins and movement behavior.
///
/// Java `ClientAnchorData extends CoordinateData`; Rust uses composition
/// because the inner type is `Copy`/`Default` and we avoid the inheritance
/// bookkeeping penalty. The four pixel margin fields match Java exactly.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ClientAnchorData {
    coordinates: CoordinateData,
    top: Option<u32>,
    right: Option<u32>,
    bottom: Option<u32>,
    left: Option<u32>,
    anchor_type: Option<AnchorType>,
}

impl ClientAnchorData {
    /// Creates a default anchor for the decorated cell. (Java default constructor)
    #[must_use]
    pub const fn new() -> Self {
        Self {
            coordinates: CoordinateData::new(),
            top: None,
            right: None,
            bottom: None,
            left: None,
            anchor_type: None,
        }
    }

    /// Sets its absolute and relative cell coordinates.
    #[must_use]
    pub const fn coordinates(mut self, value: CoordinateData) -> Self {
        self.coordinates = value;
        self
    }

    /// Sets the top margin in pixels.
    #[must_use]
    pub const fn top(mut self, value: u32) -> Self {
        self.top = Some(value);
        self
    }

    /// Sets the right margin in pixels.
    #[must_use]
    pub const fn right(mut self, value: u32) -> Self {
        self.right = Some(value);
        self
    }

    /// Sets the bottom margin in pixels.
    #[must_use]
    pub const fn bottom(mut self, value: u32) -> Self {
        self.bottom = Some(value);
        self
    }

    /// Sets the left margin in pixels.
    #[must_use]
    pub const fn left(mut self, value: u32) -> Self {
        self.left = Some(value);
        self
    }

    /// Sets the object movement and resize behavior.
    #[must_use]
    pub const fn anchor_type(mut self, value: AnchorType) -> Self {
        self.anchor_type = Some(value);
        self
    }

    /// Returns the coordinates. (Java `getCoordinates()`)
    #[must_use]
    pub const fn get_coordinates(self) -> CoordinateData {
        self.coordinates
    }

    /// Returns the top margin in pixels. (Java `getTop()`)
    #[must_use]
    pub const fn get_top(self) -> Option<u32> {
        self.top
    }

    /// Returns the right margin in pixels. (Java `getRight()`)
    #[must_use]
    pub const fn get_right(self) -> Option<u32> {
        self.right
    }

    /// Returns the bottom margin in pixels. (Java `getBottom()`)
    #[must_use]
    pub const fn get_bottom(self) -> Option<u32> {
        self.bottom
    }

    /// Returns the left margin in pixels. (Java `getLeft()`)
    #[must_use]
    pub const fn get_left(self) -> Option<u32> {
        self.left
    }

    /// Returns the movement and resize behavior. (Java `getAnchorType()`)
    #[must_use]
    pub const fn get_anchor_type(self) -> Option<AnchorType> {
        self.anchor_type
    }
}
