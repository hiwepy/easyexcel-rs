//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteHolder` (interface).

use easyexcel_core::ExcelWriteMetadata;

/// Mirrors Java `WriteHolder extends ConfigurationHolder`.
///
/// The Java interface declares seven methods; Rust mirrors only the ones a
/// handler implementation needs, because `rust_xlsxwriter` handles the rest
/// internally.
pub trait WriteHolder {
    /// Returns the resolved `ExcelWriteHeadProperty` for the holder. (Java `excelWriteHeadProperty()`)
    fn excel_write_head_property(&self) -> &ExcelWriteMetadata;

    /// Returns whether a field is ignored for the holder. (Java `ignore(fieldName, columnIndex)`)
    fn ignore(&self, field_name: Option<&str>, column_index: Option<usize>) -> bool;

    /// Returns whether a header is required. (Java `needHead()`)
    fn need_head(&self) -> bool;

    /// Returns the relative head row index. (Java `relativeHeadRowIndex()`)
    fn relative_head_row_index(&self) -> i32;

    /// Returns whether headers are auto-merged. (Java `automaticMergeHead()`)
    fn automatic_merge_head(&self) -> bool;
}