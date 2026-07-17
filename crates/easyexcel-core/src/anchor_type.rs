//! Mirrors Java `com.alibaba.excel.metadata.data.ClientAnchorData.AnchorType`.

/// Java `ClientAnchorData.AnchorType` equivalent.
///
/// Variant names are normalised to PascalCase while preserving the four POI
/// anchor modes.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AnchorType {
    /// Move and resize with the anchor cells.
    #[default]
    MoveAndResize,
    /// POI's completeness-only mode; XLSX serializes it as a one-cell anchor.
    DontMoveDoResize,
    /// Move with cells without resizing.
    MoveDontResize,
    /// Do not move or resize with cells.
    DontMoveAndResize,
}
