//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteHolder` (interface).

use std::collections::HashSet;

use crate::ExcelWriteHeadProperty;

/// Mirrors Java `WriteHolder extends ConfigurationHolder`.
pub trait WriteHolder {
    /// Returns the resolved `ExcelWriteHeadProperty` for the holder. (Java `excelWriteHeadProperty()`)
    fn excel_write_head_property(&self) -> &ExcelWriteHeadProperty;

    /// Returns whether a field is ignored for the holder. (Java `ignore(fieldName, columnIndex)`)
    fn ignore(&self, field_name: Option<&str>, column_index: Option<usize>) -> bool;

    /// Returns whether a header is required. (Java `needHead()`)
    fn need_head(&self) -> bool;

    /// Returns the relative head row index. (Java `relativeHeadRowIndex()`)
    fn relative_head_row_index(&self) -> i32;

    /// Returns whether headers are auto-merged. (Java `automaticMergeHead()`)
    fn automatic_merge_head(&self) -> bool;

    /// Returns whether output columns follow include-list order.
    /// (Java `orderByIncludeColumn()`)
    fn order_by_include_column(&self) -> bool;

    /// Returns included physical column indexes. (Java `includeColumnIndexes()`)
    fn include_column_indexes(&self) -> Option<&HashSet<usize>>;

    /// Returns included field names. (Java `includeColumnFieldNames()`)
    fn include_column_field_names(&self) -> Option<&HashSet<String>>;

    /// Returns excluded physical column indexes. (Java `excludeColumnIndexes()`)
    fn exclude_column_indexes(&self) -> Option<&HashSet<usize>>;

    /// Returns excluded field names. (Java `excludeColumnFieldNames()`)
    fn exclude_column_field_names(&self) -> Option<&HashSet<String>>;
}
