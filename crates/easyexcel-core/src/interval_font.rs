//! Mirrors Java `com.alibaba.excel.metadata.data.RichTextStringData.IntervalFont`.

use crate::write_font::WriteFont;

/// One Java `RichTextStringData.IntervalFont` range using UTF-16 indices.
///
/// Java keeps `Integer` for both indices; Rust uses `usize` to match
/// `std::str::encode_utf16` and to align with how the rest of the
/// `easyexcel-rust` workspace indexes strings.
#[derive(Debug, Clone, PartialEq)]
pub struct IntervalFont {
    start_index: usize,
    end_index: usize,
    write_font: WriteFont,
}

impl IntervalFont {
    /// Creates a half-open font range `[start_index, end_index)`. (Java inner `IntervalFont(int, int, WriteFont)`)
    #[must_use]
    pub const fn new(start_index: usize, end_index: usize, write_font: WriteFont) -> Self {
        Self {
            start_index,
            end_index,
            write_font,
        }
    }

    /// Returns the inclusive UTF-16 start index. (Java `getStartIndex()`)
    #[must_use]
    pub const fn start_index(&self) -> usize {
        self.start_index
    }

    /// Returns the exclusive UTF-16 end index. (Java `getEndIndex()`)
    #[must_use]
    pub const fn end_index(&self) -> usize {
        self.end_index
    }

    /// Returns the interval font. (Java `getWriteFont()`)
    #[must_use]
    pub const fn write_font(&self) -> &WriteFont {
        &self.write_font
    }
}
