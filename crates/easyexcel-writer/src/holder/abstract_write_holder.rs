//! Mirrors Java `com.alibaba.excel.write.metadata.holder.AbstractWriteHolder`.

use easyexcel_core::ExcelCellStyle;
use easyexcel_core::ExcelFontStyle;
use easyexcel_core::ExcelWriteMetadata;

use crate::WriteHolder;

/// Mirrors Java `AbstractWriteHolder extends AbstractHolder implements WriteHolder`.
///
/// The Java side carries type-level style, font, include/exclude, and
/// converter maps. Rust mirrors these fields with a `Copy` struct so the
/// proc-macro generated `ExcelColumn` can carry a `&'static` reference.
#[derive(Debug, Clone, Copy, Default)]
pub struct AbstractWriteHolder {
    /// Mirrors `AbstractWriteHolder.needHead`.
    pub need_head: bool,
    /// Mirrors `AbstractWriteHolder.relativeHeadRowIndex`.
    pub relative_head_row_index: i32,
    /// Mirrors `AbstractWriteHolder.useDefaultStyle`.
    pub use_default_style: bool,
    /// Mirrors `AbstractWriteHolder.automaticMergeHead`.
    pub automatic_merge_head: bool,
    /// Mirrors `AbstractWriteHolder.excelWriteHeadProperty`.
    pub excel_write_head_property: Option<ExcelWriteMetadata>,
    /// Mirrors `AbstractWriteHolder.headStyle`.
    pub head_style: Option<ExcelCellStyle>,
    /// Mirrors `AbstractWriteHolder.contentStyle`.
    pub content_style: Option<ExcelCellStyle>,
    /// Mirrors `AbstractWriteHolder.headFontStyle`.
    pub head_font_style: Option<ExcelFontStyle>,
    /// Mirrors `AbstractWriteHolder.contentFontStyle`.
    pub content_font_style: Option<ExcelFontStyle>,
}

impl WriteHolder for AbstractWriteHolder {
    fn excel_write_head_property(&self) -> &ExcelWriteMetadata {
        static EMPTY: ExcelWriteMetadata = ExcelWriteMetadata::new();
        self.excel_write_head_property.as_ref().unwrap_or(&EMPTY)
    }

    fn ignore(&self, _field_name: Option<&str>, _column_index: Option<usize>) -> bool {
        // Java `include*` / `exclude*` filtering is handled inside
        // `rust_xlsxwriter` and the derive macro; this default returns
        // `false` so handlers emit every column.
        false
    }

    fn need_head(&self) -> bool {
        self.need_head
    }

    fn relative_head_row_index(&self) -> i32 {
        self.relative_head_row_index
    }

    fn automatic_merge_head(&self) -> bool {
        self.automatic_merge_head
    }
}