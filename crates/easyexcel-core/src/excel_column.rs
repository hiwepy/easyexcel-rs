//! Mirrors Java `com.alibaba.excel.metadata.property.ExcelHeadProperty` field
//! `Head.columnIndex` / `field` / `fieldName` / `headNameList` /
//! `columnWidthProperty` / `headStyleProperty` / etc.

use crate::excel_cell_style::ExcelCellStyle;
use crate::excel_font_style::ExcelFontStyle;
use crate::metadata::property::LoopMergeProperty;

/// Static metadata for one Rust struct field and Excel column.
///
/// Mirrors the union of fields that Java stores across
/// `Head` / `FieldCache` / `FieldWrapper`. The Rust port exposes a single
/// `Copy` struct so `#[derive(ExcelRow)]` can emit a `&'static [ExcelColumn]`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ExcelColumn {
    /// Rust field name. (Java `Head.fieldName`)
    pub field: &'static str,
    /// Excel header name. (Java `Head.headNameList[0]`)
    pub name: &'static str,
    /// Explicit zero-based column index. (Java `Head.forceIndex` + `index`)
    pub index: Option<usize>,
    /// Relative ordering when no explicit index is configured. (Java `@ExcelProperty.order`)
    pub order: i32,
    /// Optional date or number format. (Java `@ExcelProperty.format`)
    pub format: Option<&'static str>,
    /// Optional annotation-driven column width in Excel character units. (Java `ColumnWidth`)
    pub column_width: Option<u16>,
    /// Field-level header cell style. (Java `@HeadStyle`)
    pub head_style: Option<ExcelCellStyle>,
    /// Field-level content cell style. (Java `@ContentStyle`)
    pub content_style: Option<ExcelCellStyle>,
    /// Field-level header font style. (Java `@HeadFontStyle`)
    pub head_font_style: Option<ExcelFontStyle>,
    /// Field-level content font style. (Java `@ContentFontStyle`)
    pub content_font_style: Option<ExcelFontStyle>,
    /// Field-level repeating content merge. (Java `@ContentLoopMerge` → `Head.loopMergeProperty`)
    pub loop_merge: Option<LoopMergeProperty>,
}

impl ExcelColumn {
    /// Creates static column metadata. (Java `Head(columnIndex, field, fieldName, headNameList, forceIndex, forceName)` subset)
    #[allow(clippy::too_many_arguments)]
    #[must_use]
    pub const fn new(
        field: &'static str,
        name: &'static str,
        index: Option<usize>,
        order: i32,
        format: Option<&'static str>,
    ) -> Self {
        Self {
            field,
            name,
            index,
            order,
            format,
            column_width: None,
            head_style: None,
            content_style: None,
            head_font_style: None,
            content_font_style: None,
            loop_merge: None,
        }
    }

    /// Adds annotation-driven column width. (Java `@ColumnWidth`)
    #[must_use]
    pub const fn with_column_width(mut self, width: u16) -> Self {
        self.column_width = Some(width);
        self
    }

    /// Adds a field-level header cell style. (Java `@HeadStyle`)
    #[must_use]
    pub const fn with_head_style(mut self, style: ExcelCellStyle) -> Self {
        self.head_style = Some(style);
        self
    }

    /// Adds a field-level content cell style. (Java `@ContentStyle`)
    #[must_use]
    pub const fn with_content_style(mut self, style: ExcelCellStyle) -> Self {
        self.content_style = Some(style);
        self
    }

    /// Adds a field-level header font style. (Java `@HeadFontStyle`)
    #[must_use]
    pub const fn with_head_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.head_font_style = Some(style);
        self
    }

    /// Adds a field-level content font style. (Java `@ContentFontStyle`)
    #[must_use]
    pub const fn with_content_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.content_font_style = Some(style);
        self
    }

    /// Adds a field-level repeating content merge. (Java `@ContentLoopMerge`)
    #[must_use]
    pub const fn with_loop_merge(mut self, property: LoopMergeProperty) -> Self {
        self.loop_merge = Some(property);
        self
    }
}
