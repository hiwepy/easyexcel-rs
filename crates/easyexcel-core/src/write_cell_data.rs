//! Mirrors Java `com.alibaba.excel.metadata.data.WriteCellData`.

use crate::ExcelCellStyle;
use crate::cell_value::CellValue;
use crate::comment_data::CommentData;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;
use crate::formula_data::FormulaData;
use crate::from_excel_cell::FromExcelCell;
use crate::hyperlink_data::HyperlinkData;
use crate::image_data::ImageData;
use crate::into_excel_cell::IntoExcelCell;
use crate::metadata::data::DataFormatData;
use crate::rich_text_string_data::RichTextStringData;

/// Java `WriteCellData` subset that preserves a scalar plus decorations.
///
/// Java `WriteCellData` extends `CellData` and adds image / comment / hyperlink
/// / formula fields. Rust keeps the same public surface on the hot path.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellData {
    value: CellValue,
    image_data_list: Vec<ImageData>,
    comment_data: Option<CommentData>,
    hyperlink_data: Option<HyperlinkData>,
    formula_data: Option<FormulaData>,
    write_cell_style: Option<ExcelCellStyle>,
    data_format_data: Option<DataFormatData>,
}

impl WriteCellData {
    /// Creates decorated cell data from a scalar value. (Java `WriteCellData(WriteCellData)`)
    #[must_use]
    pub const fn new(value: CellValue) -> Self {
        Self {
            value,
            image_data_list: Vec::new(),
            comment_data: None,
            hyperlink_data: None,
            formula_data: None,
            write_cell_style: None,
            data_format_data: None,
        }
    }

    /// Creates a string cell. (Java `WriteCellData(String)`)
    #[must_use]
    pub fn from_string(value: impl Into<String>) -> Self {
        Self::new(CellValue::String(value.into()))
    }

    /// Creates an empty scalar cell with one image, matching Java's byte-array constructor.
    #[must_use]
    pub fn from_image(image: impl Into<Vec<u8>>) -> Self {
        Self::new(CellValue::Empty).image(ImageData::new(image))
    }

    /// Creates a rich-text cell, matching Java's `RICH_TEXT_STRING` cell data type.
    #[must_use]
    pub const fn from_rich_text(value: RichTextStringData) -> Self {
        Self::new(CellValue::RichText(value))
    }

    /// Creates a hyperlink cell with optional display text. (Java `WriteCellData.setHyperlinkData(...)`)
    #[must_use]
    pub fn from_hyperlink(url: impl Into<String>, text: impl Into<String>) -> Self {
        Self::new(CellValue::Hyperlink {
            url: url.into(),
            text: text.into(),
        })
    }

    /// Creates a formula cell. (Java `WriteCellData.setFormulaData(...)`)
    #[must_use]
    pub fn from_formula(formula: impl Into<String>) -> Self {
        Self::new(CellValue::Formula(formula.into()))
    }

    /// Creates a comment-decorated cell. (Java `WriteCellData.setCommentData(...)`)
    #[must_use]
    pub fn from_comment(value: impl Into<CellValue>, text: impl Into<String>) -> Self {
        Self::new(CellValue::Comment {
            value: Box::new(value.into()),
            text: text.into(),
        })
    }

    /// Replaces the underlying scalar value while keeping decorations intact.
    ///
    /// Mirrors Java's `WriteCellData.setValue(...)` setter used by the writer
    /// when an annotation override (formula / hyperlink) needs to wrap the
    /// typed scalar without reallocating the cell structure.
    pub fn set_value(&mut self, value: impl Into<CellValue>) -> &mut Self {
        self.value = value.into();
        self
    }

    /// Appends one image entry. (Java `setImageDataList(List<ImageData>)` step)
    #[must_use]
    pub fn image(mut self, value: ImageData) -> Self {
        self.image_data_list.push(value);
        self
    }

    /// Replaces the full image list.
    #[must_use]
    pub fn image_data_list(mut self, value: impl IntoIterator<Item = ImageData>) -> Self {
        self.image_data_list = value.into_iter().collect();
        self
    }

    /// Sets comment metadata. (Java `setCommentData(CommentData)`)
    #[must_use]
    pub fn comment_data(mut self, value: CommentData) -> Self {
        self.comment_data = Some(value);
        self
    }

    /// Sets hyperlink metadata. (Java `setHyperlinkData(HyperlinkData)`)
    #[must_use]
    pub fn hyperlink_data(mut self, value: HyperlinkData) -> Self {
        self.hyperlink_data = Some(value);
        self
    }

    /// Sets formula metadata. (Java `setFormulaData(FormulaData)`)
    #[must_use]
    pub fn formula_data(mut self, value: FormulaData) -> Self {
        self.formula_data = Some(value);
        self
    }

    /// Returns the scalar cell value. (Java `getValue()` via `CellData.getData()`)
    #[must_use]
    pub const fn value(&self) -> &CellValue {
        &self.value
    }

    /// Returns all image entries in insertion order. (Java `getImageDataList()`)
    #[must_use]
    pub fn images(&self) -> &[ImageData] {
        &self.image_data_list
    }

    /// Returns comment metadata. (Java `getCommentData()`)
    #[must_use]
    pub const fn get_comment_data(&self) -> Option<&CommentData> {
        self.comment_data.as_ref()
    }

    /// Returns hyperlink metadata. (Java `getHyperlinkData()`)
    #[must_use]
    pub const fn get_hyperlink_data(&self) -> Option<&HyperlinkData> {
        self.hyperlink_data.as_ref()
    }

    /// Returns formula metadata. (Java `getFormulaData()`)
    #[must_use]
    pub const fn get_formula_data(&self) -> Option<&FormulaData> {
        self.formula_data.as_ref()
    }

    /// Returns the logical cell style. (Java `getWriteCellStyle()`)
    #[must_use]
    pub const fn write_cell_style(&self) -> Option<&ExcelCellStyle> {
        self.write_cell_style.as_ref()
    }

    /// Replaces the logical cell style. (Java `setWriteCellStyle(...)`)
    pub fn set_write_cell_style(&mut self, style: Option<ExcelCellStyle>) {
        self.write_cell_style = style;
    }

    /// Returns a mutable style, creating it when absent.
    ///
    /// Mirrors Java `WriteCellData#getOrCreateStyle`.
    pub fn get_or_create_style(&mut self) -> &mut ExcelCellStyle {
        self.write_cell_style
            .get_or_insert_with(ExcelCellStyle::default)
    }

    /// Returns the owned data-format metadata associated with the style.
    ///
    /// Java stores this object inside `WriteCellStyle`; Rust keeps the owned
    /// runtime string beside the copyable annotation style.
    #[must_use]
    pub const fn data_format_data(&self) -> Option<&DataFormatData> {
        self.data_format_data.as_ref()
    }

    /// Returns mutable data-format metadata, creating it when absent.
    pub fn get_or_create_data_format(&mut self) -> &mut DataFormatData {
        self.data_format_data
            .get_or_insert_with(DataFormatData::default)
    }

    /// Resolves the scalar plus formula/link/comment/image decorations into
    /// the backend-neutral value written by an engine.
    ///
    /// The style and data-format fields intentionally remain on
    /// `WriteCellData`; Java applies them after conversion in
    /// `FillStyleCellWriteHandler`.
    #[must_use]
    pub fn effective_value(&self) -> CellValue {
        let mut value = self.value.clone();
        if let Some(formula) = &self.formula_data {
            value = CellValue::Formula(formula.formula_value().to_owned());
        }
        if let Some(link) = &self.hyperlink_data {
            let url = link.get_address().unwrap_or("").to_owned();
            let text = match &value {
                CellValue::String(s) => s.clone(),
                other => other.as_text(),
            };
            value = CellValue::Hyperlink { url, text };
        }
        if let Some(comment) = &self.comment_data {
            value = CellValue::Comment {
                value: Box::new(value),
                text: comment.note_text(),
            };
        }
        if self.image_data_list.is_empty() {
            value
        } else {
            CellValue::Images {
                value: Box::new(value),
                images: self.image_data_list.clone(),
            }
        }
    }
}

impl IntoExcelCell for WriteCellData {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(self.effective_value())
    }
}

impl FromExcelCell for WriteCellData {
    fn from_excel_cell(
        cell: Option<&CellValue>,
        _context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        Ok(Self::new(cell.cloned().unwrap_or(CellValue::Empty)))
    }
}
