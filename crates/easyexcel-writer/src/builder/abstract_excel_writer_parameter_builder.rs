//! Mirrors Java `com.alibaba.excel.write.builder.AbstractExcelWriterParameterBuilder`.

use crate::CellStyle;
use easyexcel_core::WriteHandler;

use crate::metadata::WriteBasicParameter;

/// Mirrors Java `AbstractExcelWriterParameterBuilder<T, C>`.
///
/// The Java side chains 12 setter methods (`needHead`, `useDefaultStyle`,
/// `automaticMergeHead`, `excludeColumnIndexes`, `excludeColumnFieldNames`,
/// `includeColumnIndexes`, `includeColumnFieldNames`,
/// `orderByIncludeColumn`, `relativeHeadRowIndex`, `registerWriteHandler`,
/// `excludeColumnFiledNames` (typo'd alias), and `head(List)`).
///
/// In Rust, the same surface lives on the chain-returning
/// [`crate::EasyExcel::write`]-based builder. This trait preserves the
/// 1:1 names so Java-aware code can still find the canonical setters.
pub trait AbstractExcelWriterParameterBuilder {
    /// Returns the parameter being mutated. (Java `parameter()`)
    fn parameter(&mut self) -> &mut WriteBasicParameter;

    /// Sets whether a header row is written. (Java `needHead(Boolean)`)
    fn need_head(&mut self, need_head: bool) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().need_head = Some(need_head);
        self
    }

    /// Sets the default style flag. (Java `useDefaultStyle(Boolean)`)
    fn use_default_style(&mut self, use_default_style: bool) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().use_default_style = Some(use_default_style);
        self
    }

    /// Sets automatic header merging. (Java `automaticMergeHead(Boolean)`)
    fn automatic_merge_head(&mut self, automatic_merge_head: bool) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().automatic_merge_head = Some(automatic_merge_head);
        self
    }

    /// Sets the relative head row index. (Java `relativeHeadRowIndex(Integer)`)
    fn relative_head_row_index(&mut self, index: i32) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().relative_head_row_index = Some(index);
        self
    }

    /// Sets the include-order flag. (Java `orderByIncludeColumn(Boolean)`)
    fn order_by_include_column(&mut self, enabled: bool) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().order_by_include_column = Some(enabled);
        self
    }

    /// Replaces inherited excluded physical columns.
    /// (Java `excludeColumnIndexes(Collection<Integer>)`)
    fn exclude_column_indexes(&mut self, indexes: Vec<usize>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().exclude_column_indexes = Some(indexes);
        self
    }

    /// Replaces inherited excluded field names.
    /// (Java `excludeColumnFieldNames(Collection<String>)`)
    fn exclude_column_field_names(&mut self, names: Vec<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().exclude_column_field_names = Some(names);
        self
    }

    /// Deprecated Java spelling retained for migration compatibility.
    fn exclude_column_filed_names(&mut self, names: Vec<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.exclude_column_field_names(names)
    }

    /// Replaces inherited included physical columns.
    /// (Java `includeColumnIndexes(Collection<Integer>)`)
    fn include_column_indexes(&mut self, indexes: Vec<usize>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().include_column_indexes = Some(indexes);
        self
    }

    /// Replaces inherited included field names.
    /// (Java `includeColumnFieldNames(Collection<String>)`)
    fn include_column_field_names(&mut self, names: Vec<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().include_column_field_names = Some(names);
        self
    }

    /// Deprecated Java spelling retained for migration compatibility.
    fn include_column_filed_names(&mut self, names: Vec<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.include_column_field_names(names)
    }

    /// Appends a write handler. (Java `registerWriteHandler(WriteHandler)`)
    fn register_write_handler(&mut self, handler: Box<dyn WriteHandler>) -> &mut Self
    where
        Self: Sized;

    /// Convenience setter that returns a `CellStyle` to builder methods. The
    /// Java side exposes typed setters on `ExcelWriterBuilder.head_style`; the
    /// trait accepts the value object directly.
    fn head_style_slot(&self) -> Option<CellStyle> {
        None
    }
}
