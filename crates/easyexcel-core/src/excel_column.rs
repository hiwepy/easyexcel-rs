//! Mirrors Java `com.alibaba.excel.metadata.property.ExcelHeadProperty` field
//! `Head.columnIndex` / `field` / `fieldName` / `headNameList` /
//! `columnWidthProperty` / `headStyleProperty` / etc.

use crate::cell_value::CellValue;
use crate::comment_data::CommentData;
use crate::excel_cell_style::ExcelCellStyle;
use crate::excel_font_style::ExcelFontStyle;
use crate::hyperlink_data::HyperlinkData;
use crate::metadata::property::data_validation_property::ExcelDataValidationMeta;
use crate::metadata::property::LoopMergeProperty;
use crate::write_cell_data::WriteCellData;

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

    // Phase 1: new annotation-derived fields (Phase 1 markers in
    // com.alibaba.excel.annotation.write.*ExcelImage / Comment / Hyperlink /
    // Formula / DataValidation / Conditional / Filter).
    /// Optional image path or URL for this column. (Java `@ExcelImage.image()`)
    pub image_path: Option<&'static str>,
    /// Optional cell comment / note. (Java `@ExcelComment.value()`)
    pub comment: Option<&'static str>,
    /// Optional hyperlink target. (Java `@ExcelHyperlink.value()`)
    pub hyperlink: Option<&'static str>,
    /// Optional formula override. (Java `@ExcelFormula.value()`)
    pub formula: Option<&'static str>,
    /// Optional data-validation metadata. (Java `@ExcelDataValidation`)
    pub data_validation: Option<ExcelDataValidationMeta>,
    /// Optional conditional-formatting tuple `(condition, font_color, bg_color)`.
    /// (Java `@ExcelConditional`)
    pub conditional_format: Option<(&'static str, &'static str, &'static str)>,
    /// Whether this column participates in auto-filtering. (Java `@ExcelFilter`)
    pub auto_filter: bool,
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
            image_path: None,
            comment: None,
            hyperlink: None,
            formula: None,
            data_validation: None,
            conditional_format: None,
            auto_filter: false,
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

    // Phase 1: new annotation-derived column builders

    /// Adds a per-column image source. (Java `@ExcelImage`)
    #[must_use]
    pub const fn with_image_path(mut self, path: &'static str) -> Self {
        self.image_path = Some(path);
        self
    }

    /// Adds a per-column cell comment. (Java `@ExcelComment`)
    #[must_use]
    pub const fn with_comment(mut self, comment: &'static str) -> Self {
        self.comment = Some(comment);
        self
    }

    /// Adds a per-column hyperlink target. (Java `@ExcelHyperlink`)
    #[must_use]
    pub const fn with_hyperlink(mut self, link: &'static str) -> Self {
        self.hyperlink = Some(link);
        self
    }

    /// Adds a per-column formula override. (Java `@ExcelFormula`)
    #[must_use]
    pub const fn with_formula(mut self, formula: &'static str) -> Self {
        self.formula = Some(formula);
        self
    }

    /// Adds per-column data-validation metadata. (Java `@ExcelDataValidation`)
    #[must_use]
    pub const fn with_data_validation(mut self, meta: ExcelDataValidationMeta) -> Self {
        self.data_validation = Some(meta);
        self
    }

    /// Adds per-column conditional-formatting metadata. (Java `@ExcelConditional`)
    #[must_use]
    pub const fn with_conditional_format(
        mut self,
        cf: (&'static str, &'static str, &'static str),
    ) -> Self {
        self.conditional_format = Some(cf);
        self
    }

    /// Marks the column as participating in auto-filter. (Java `@ExcelFilter`)
    #[must_use]
    pub const fn with_auto_filter(mut self) -> Self {
        self.auto_filter = true;
        self
    }

    // -------- Phase 1.4: decoration helpers applied to WriteCellData --------

    /// Applies this column's annotation-driven decorations (hyperlink / formula /
    /// comment) onto a `WriteCellData`. (Java `Head.fillHeadAndWriteData` decorations)
    ///
    /// Order matches Java `ExcelBuilderImpl` write path:
    /// 1. formula override wraps the scalar (`CellValue::Formula`)
    /// 2. hyperlink wraps the display text (`CellValue::Hyperlink`)
    /// 3. comment wraps the underlying value (`CellValue::Comment`)
    pub fn apply_decorations(&self, mut data: WriteCellData) -> WriteCellData {
        if let Some(formula) = self.formula {
            data.set_value(CellValue::Formula(formula.to_owned()));
        }
        if let Some(url) = self.hyperlink {
            let text = match data.value() {
                CellValue::String(s) => s.clone(),
                other => other.as_text(),
            };
            data.set_value(CellValue::Hyperlink {
                url: url.to_owned(),
                text,
            });
            // Also reflect via HyperlinkData so writer layers can access
            // both the wrapped CellValue and the structured side-channel.
            data = data.hyperlink_data(HyperlinkData::new().address(url.to_owned()));
        }
        if let Some(comment_text) = self.comment {
            let comment = CommentData::new().text(comment_text.to_owned());
            data = data.comment_data(comment);
        }
        data
    }
}