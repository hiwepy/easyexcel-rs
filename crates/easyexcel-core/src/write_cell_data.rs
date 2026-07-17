//! Mirrors Java `com.alibaba.excel.metadata.data.WriteCellData`.

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::excel_error::ExcelError;
use crate::from_excel_cell::FromExcelCell;
use crate::image_data::ImageData;
use crate::into_excel_cell::IntoExcelCell;
use crate::rich_text_string_data::RichTextStringData;

/// Java `WriteCellData` subset that preserves a scalar plus `imageDataList`.
///
/// Java `WriteCellData` extends `CellData` and adds image / comment / hyperlink
/// fields. Rust collapses the scalar + image list into a single `CellValue`
/// variant, mirroring the Java fields and constructors on the hot path.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellData {
    value: CellValue,
    image_data_list: Vec<ImageData>,
}

impl WriteCellData {
    /// Creates decorated cell data from a scalar value. (Java `WriteCellData(WriteCellData)`)
    #[must_use]
    pub const fn new(value: CellValue) -> Self {
        Self {
            value,
            image_data_list: Vec::new(),
        }
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
}

impl IntoExcelCell for WriteCellData {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Images {
            value: Box::new(self.value.clone()),
            images: self.image_data_list.clone(),
        })
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
