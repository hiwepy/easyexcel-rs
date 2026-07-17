//! Mirrors Java `com.alibaba.excel.metadata.AbstractHolder` type-level fields
//! aggregated into `WriteBasicParameter` plus the
//! `ExcelWriteHeadProperty` derivation.

use crate::excel_cell_style::ExcelCellStyle;
use crate::excel_font_style::ExcelFontStyle;

/// Type-level dimensions derived from Java-style write annotations.
///
/// Java stores these fields on `AbstractWriteHolder`. Rust emits a single
/// `Copy` struct from `#[derive(ExcelRow)]`.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ExcelWriteMetadata {
    /// Default width for columns without a field-level override. (Java `@ColumnWidth` on type)
    pub column_width: Option<u16>,
    /// Height for every generated header row. (Java `@HeadRowHeight`)
    pub head_row_height: Option<u16>,
    /// Height for every generated content row. (Java `@ContentRowHeight`)
    pub content_row_height: Option<u16>,
    /// Type-level header cell style. (Java `@HeadStyle` on type)
    pub head_style: Option<ExcelCellStyle>,
    /// Type-level content cell style. (Java `@ContentStyle` on type)
    pub content_style: Option<ExcelCellStyle>,
    /// Type-level header font style. (Java `@HeadFontStyle` on type)
    pub head_font_style: Option<ExcelFontStyle>,
    /// Type-level content font style. (Java `@ContentFontStyle` on type)
    pub content_font_style: Option<ExcelFontStyle>,
}

impl ExcelWriteMetadata {
    /// Creates empty write metadata.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            column_width: None,
            head_row_height: None,
            content_row_height: None,
            head_style: None,
            content_style: None,
            head_font_style: None,
            content_font_style: None,
        }
    }

    /// Sets the type-level default column width. (Java `AbstractWriteHolder.columnWidth`)
    #[must_use]
    pub const fn column_width(mut self, width: u16) -> Self {
        self.column_width = Some(width);
        self
    }

    /// Sets the generated header-row height. (Java `AbstractWriteHolder.headRowHeight`)
    #[must_use]
    pub const fn head_row_height(mut self, height: u16) -> Self {
        self.head_row_height = Some(height);
        self
    }

    /// Sets the generated content-row height. (Java `AbstractWriteHolder.contentRowHeight`)
    #[must_use]
    pub const fn content_row_height(mut self, height: u16) -> Self {
        self.content_row_height = Some(height);
        self
    }

    /// Adds a type-level header cell style. (Java `AbstractWriteHolder.headStyle`)
    #[must_use]
    pub const fn head_style(mut self, style: ExcelCellStyle) -> Self {
        self.head_style = Some(style);
        self
    }

    /// Adds a type-level content cell style. (Java `AbstractWriteHolder.contentStyle`)
    #[must_use]
    pub const fn content_style(mut self, style: ExcelCellStyle) -> Self {
        self.content_style = Some(style);
        self
    }

    /// Adds a type-level header font style. (Java `AbstractWriteHolder.headFontStyle`)
    #[must_use]
    pub const fn head_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.head_font_style = Some(style);
        self
    }

    /// Adds a type-level content font style. (Java `AbstractWriteHolder.contentFontStyle`)
    #[must_use]
    pub const fn content_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.content_font_style = Some(style);
        self
    }
}
