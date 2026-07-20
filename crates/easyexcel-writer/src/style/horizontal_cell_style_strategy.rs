//! Mirrors Java `com.alibaba.excel.write.style.HorizontalCellStyleStrategy`.
//!
//! Wired into the XLSX write path via [`WriteHandler::style_cell_style`], which
//! the writer merges into each cell format after annotation styles. Nested
//! fonts on [`ExcelCellStyle::font`] mirror Java `WriteCellStyle.writeFont`.

use easyexcel_core::{ExcelCellStyle, ExcelFontStyle, WriteCellContext, WriteFont, WriteHandler};

use crate::metadata::style::write_font::excel_font_style_from_write_font;
use crate::style::abstract_cell_style_strategy::AbstractCellStyleStrategy;

/// Mirrors Java `HorizontalCellStyleStrategy`.
///
/// The Java side cycles through a list of content styles by
/// `relativeRowIndex`; the Rust port mirrors that behaviour once the
/// write path supplies [`WriteCellContext::relative_row_index`].
/// Styles may carry nested fonts via [`ExcelCellStyle::font`] (Java
/// `WriteCellStyle.setWriteFont`).
pub struct HorizontalCellStyleStrategy {
    head_style: ExcelCellStyle,
    content_styles: Vec<ExcelCellStyle>,
}

impl HorizontalCellStyleStrategy {
    /// Creates a strategy with content styles only (empty head style).
    /// (Java `HorizontalCellStyleStrategy(List<WriteCellStyle>)` subset)
    #[must_use]
    pub const fn new(content_styles: Vec<ExcelCellStyle>) -> Self {
        Self {
            head_style: ExcelCellStyle::new(),
            content_styles,
        }
    }

    /// Creates a strategy with one head style and one content style.
    /// (Java `HorizontalCellStyleStrategy(WriteCellStyle, WriteCellStyle)`)
    #[must_use]
    pub fn with_head_and_content(
        head_style: ExcelCellStyle,
        content_style: ExcelCellStyle,
    ) -> Self {
        Self {
            head_style,
            content_styles: vec![content_style],
        }
    }

    /// Creates a strategy with one head style and a content-style cycle.
    /// (Java `HorizontalCellStyleStrategy(WriteCellStyle, List<WriteCellStyle>)`)
    #[must_use]
    pub const fn with_head_and_contents(
        head_style: ExcelCellStyle,
        content_styles: Vec<ExcelCellStyle>,
    ) -> Self {
        Self {
            head_style,
            content_styles,
        }
    }

    /// Attaches a head font (Java `headWriteCellStyle.setWriteFont`).
    #[must_use]
    pub const fn with_head_font(mut self, font: ExcelFontStyle) -> Self {
        self.head_style.font = Some(font);
        self
    }

    /// Attaches a head font from runtime [`WriteFont`]
    /// (Java `WriteCellStyle.setWriteFont(WriteFont)`).
    ///
    /// Owned font names are not copied into [`ExcelFontStyle`]; set
    /// [`ExcelFontStyle::font_name`] when a static name is required.
    #[must_use]
    pub fn with_head_write_font(mut self, font: WriteFont) -> Self {
        self.head_style.font = Some(excel_font_style_from_write_font(&font));
        self
    }

    /// Attaches one content font to every configured content style
    /// (Java each `contentWriteCellStyle.setWriteFont`).
    #[must_use]
    pub fn with_content_font(mut self, font: ExcelFontStyle) -> Self {
        for style in &mut self.content_styles {
            style.font = Some(font);
        }
        self
    }

    /// Attaches a content font from runtime [`WriteFont`].
    #[must_use]
    pub fn with_content_write_font(mut self, font: WriteFont) -> Self {
        let converted = excel_font_style_from_write_font(&font);
        for style in &mut self.content_styles {
            style.font = Some(converted);
        }
        self
    }

    /// Returns the configured head style. (Java `getHeadWriteCellStyle()`)
    #[must_use]
    pub const fn head_style(&self) -> ExcelCellStyle {
        self.head_style
    }

    /// Returns the configured content styles. (Java `getContentWriteCellStyleList()`)
    #[must_use]
    pub fn content_styles(&self) -> &[ExcelCellStyle] {
        &self.content_styles
    }
}

impl AbstractCellStyleStrategy for HorizontalCellStyleStrategy {
    fn cell_style(&self, context: &WriteCellContext) -> ExcelCellStyle {
        // Java `setHeadCellStyle` / `setContentCellStyle`
        if context.is_head {
            return self.head_style;
        }
        if self.content_styles.is_empty() {
            return ExcelCellStyle::new();
        }
        // Java: `relativeRowIndex % contentWriteCellStyleList.size()`
        let relative = context.relative_row_index.unwrap_or(0);
        self.content_styles[relative % self.content_styles.len()]
    }
}

impl WriteHandler for HorizontalCellStyleStrategy {
    fn order(&self) -> i32 {
        // Java `OrderConstant.DEFINE_STYLE` on `AbstractCellStyleStrategy`
        50_000
    }

    fn style_cell_style(&self, context: &WriteCellContext) -> Option<ExcelCellStyle> {
        Some(AbstractCellStyleStrategy::cell_style(self, context))
    }
}
