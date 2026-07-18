//! Mirrors Java `com.alibaba.excel.write.property.ExcelWriteHeadProperty`.

use easyexcel_core::ExcelWriteMetadata;

/// Mirrors Java `ExcelWriteHeadProperty extends ExcelHeadProperty`.
///
/// Java's type carries a `headMap: Map<Integer, Head>` plus
/// `headRowHeightProperty`, `contentRowHeightProperty`, and
/// `onceAbsoluteMergeProperty`. Rust exposes the same data through
/// [`ExcelWriteMetadata`] and a `Copy` handle so derive macro
/// can return `&'static` references.
pub struct ExcelWriteHeadProperty {
    /// Mirrors `ExcelWriteHeadProperty.headRowHeightProperty`.
    pub head_row_height: Option<u16>,
    /// Mirrors `ExcelWriteHeadProperty.contentRowHeightProperty`.
    pub content_row_height: Option<u16>,
    /// Mirrors `ExcelWriteHeadProperty.onceAbsoluteMergeProperty`.
    pub once_absolute_merge: Option<ExcelWriteMetadata>,
}

impl ExcelWriteHeadProperty {
    /// Creates an empty property.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            head_row_height: None,
            content_row_height: None,
            once_absolute_merge: None,
        }
    }

    /// Returns the head row height. (Java `getHeadRowHeightProperty()`)
    #[must_use]
    pub const fn head_row_height(&self) -> Option<u16> {
        self.head_row_height
    }

    /// Returns the content row height. (Java `getContentRowHeightProperty()`)
    #[must_use]
    pub const fn content_row_height(&self) -> Option<u16> {
        self.content_row_height
    }

    /// Returns the once-absolute merge range, if any. (Java
    /// `getOnceAbsoluteMergeProperty()`)
    #[must_use]
    pub const fn once_absolute_merge(&self) -> Option<&ExcelWriteMetadata> {
        self.once_absolute_merge.as_ref()
    }
}

impl Default for ExcelWriteHeadProperty {
    fn default() -> Self {
        Self::new()
    }
}