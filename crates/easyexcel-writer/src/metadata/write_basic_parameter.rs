//! Mirrors Java `com.alibaba.excel.write.metadata.WriteBasicParameter`.

use easyexcel_core::ConverterRegistry;

/// Mirrors Java `WriteBasicParameter extends BasicParameter`.
///
/// Java carries 9 fields (`relativeHeadRowIndex`, `needHead`,
/// `customWriteHandlerList`, `useDefaultStyle`, `automaticMergeHead`,
/// `excludeColumnIndexes`, `excludeColumnFieldNames`,
/// `includeColumnIndexes`, `includeColumnFieldNames`,
/// `orderByIncludeColumn`). Rust reuses [`WriteOptions`] for the same
/// data, and uses this struct as a thin handle so the 1:1 API name is
/// preserved.
#[derive(Debug, Clone, Default)]
pub struct WriteBasicParameter {
    /// Mirrors `WriteBasicParameter.relativeHeadRowIndex`.
    pub relative_head_row_index: i32,
    /// Mirrors `WriteBasicParameter.needHead`.
    pub need_head: bool,
    /// Mirrors `WriteBasicParameter.useDefaultStyle`.
    pub use_default_style: bool,
    /// Mirrors `WriteBasicParameter.automaticMergeHead`.
    pub automatic_merge_head: bool,
    /// Mirrors `WriteBasicParameter.orderByIncludeColumn`.
    pub order_by_include_column: bool,
    /// Mirrors `WriteBasicParameter.converters` (custom-registered converters).
    pub converters: ConverterRegistry,
}

impl WriteBasicParameter {
    /// Returns whether a header row is required. (Java `getNeedHead()`)
    #[must_use]
    pub const fn get_need_head(&self) -> bool {
        self.need_head
    }

    /// Returns the relative head row index. (Java `getRelativeHeadRowIndex()`)
    #[must_use]
    pub const fn get_relative_head_row_index(&self) -> i32 {
        self.relative_head_row_index
    }

    /// Returns whether headers are auto-merged. (Java `getAutomaticMergeHead()`)
    #[must_use]
    pub const fn get_automatic_merge_head(&self) -> bool {
        self.automatic_merge_head
    }
}