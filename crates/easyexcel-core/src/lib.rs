//! Core data model and extension points for `easyexcel-rs`.

use std::any::{Any, TypeId, type_name};
use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::fmt::Display;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::Arc;
use std::time::Duration;

/// Arbitrary-precision decimal type used for Java `BigDecimal`-compatible cells.
pub use bigdecimal::BigDecimal;
use bigdecimal::ToPrimitive;
use chrono::{NaiveDate, NaiveDateTime};
/// Arbitrary-precision integer type used for Java `BigInteger`-compatible fields.
pub use num_bigint::BigInt;
use thiserror::Error;
/// Parsed URL type accepted by the default Java-compatible URL image converter.
pub use url::Url;

/// The result type used by all easyexcel crates.
pub type Result<T> = std::result::Result<T, ExcelError>;

/// Character encoding used by the CSV reader and writer.
///
/// Names follow Java's `Charset.forName` convention. The backend accepts
/// case-insensitive WHATWG labels such as `UTF-8`, `UTF-16BE`, `GBK`, and
/// `windows-1252`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CsvCharset(String);

impl CsvCharset {
    /// Creates a charset from a Java-style charset name or alias.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self(name.into())
    }

    /// Returns the configured charset name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.0
    }

    /// Returns UTF-8, the deterministic Rust default.
    #[must_use]
    pub fn utf8() -> Self {
        Self("UTF-8".to_owned())
    }
}

impl Default for CsvCharset {
    fn default() -> Self {
        Self::utf8()
    }
}

impl From<&str> for CsvCharset {
    fn from(value: &str) -> Self {
        Self::new(value)
    }
}

impl From<String> for CsvCharset {
    fn from(value: String) -> Self {
        Self::new(value)
    }
}

/// A backend-neutral Excel cell value.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// An absent or blank cell.
    Empty,
    /// A string cell.
    String(String),
    /// A Boolean cell.
    Bool(bool),
    /// A signed integer cell.
    Int(i64),
    /// A floating-point cell.
    Float(f64),
    /// An arbitrary-precision decimal cell matching Java `BigDecimal` reads.
    Decimal(BigDecimal),
    /// A date-only cell.
    Date(NaiveDate),
    /// A date and time cell.
    DateTime(NaiveDateTime),
    /// An Excel error value.
    Error(String),
    /// An Excel formula expression without the leading `=` requirement.
    Formula(String),
    /// A clickable hyperlink and its displayed text.
    Hyperlink {
        /// Link target.
        url: String,
        /// Displayed cell text.
        text: String,
    },
    /// A value decorated with an Excel cell note/comment.
    Comment {
        /// Underlying cell value.
        value: Box<CellValue>,
        /// Note text.
        text: String,
    },
    /// Encoded PNG, JPEG, GIF, or BMP image bytes.
    Image(Vec<u8>),
    /// Text with whole-string and interval font metadata.
    RichText(RichTextStringData),
    /// A scalar cell decorated with Java-compatible `WriteCellData.imageDataList` entries.
    Images {
        /// Scalar value written before the drawings are added.
        value: Box<CellValue>,
        /// Images and their worksheet anchor metadata.
        images: Vec<ImageData>,
    },
}

impl CellValue {
    /// Returns whether the cell is empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }

    /// Returns a deterministic textual representation used for header matching.
    #[must_use]
    pub fn as_text(&self) -> String {
        match self {
            Self::Empty | Self::Image(_) => String::new(),
            Self::String(value) | Self::Error(value) | Self::Formula(value) => value.clone(),
            Self::Bool(value) => value.to_string(),
            Self::Int(value) => value.to_string(),
            Self::Float(value) => value.to_string(),
            Self::Decimal(value) => value.to_string(),
            Self::Date(value) => value.format("%Y-%m-%d").to_string(),
            Self::DateTime(value) => value.format("%Y-%m-%d %H:%M:%S").to_string(),
            Self::Hyperlink { text, .. } => text.clone(),
            Self::RichText(value) => value.text_string().to_owned(),
            Self::Comment { value, .. } | Self::Images { value, .. } => value.as_text(),
        }
    }

    /// Returns the Java EasyExcel-compatible logical cell type used to select converters.
    #[must_use]
    pub fn data_type(&self) -> CellDataType {
        match self {
            Self::Empty => CellDataType::Empty,
            Self::String(_) | Self::Hyperlink { .. } => CellDataType::String,
            Self::Bool(_) => CellDataType::Boolean,
            Self::Int(_) | Self::Float(_) | Self::Decimal(_) => CellDataType::Number,
            Self::Date(_) | Self::DateTime(_) => CellDataType::Date,
            Self::RichText(_) => CellDataType::RichTextString,
            Self::Error(_) => CellDataType::Error,
            Self::Formula(_) => CellDataType::Formula,
            Self::Comment { value, .. } | Self::Images { value, .. } => value.data_type(),
            Self::Image(_) => CellDataType::Image,
        }
    }
}

/// Cell coordinates used by Java `CoordinateData` decorations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[allow(clippy::struct_field_names)]
pub struct CoordinateData {
    first_row_index: Option<u32>,
    first_column_index: Option<u16>,
    last_row_index: Option<u32>,
    last_column_index: Option<u16>,
    relative_first_row_index: Option<i32>,
    relative_first_column_index: Option<i32>,
    relative_last_row_index: Option<i32>,
    relative_last_column_index: Option<i32>,
}

impl CoordinateData {
    /// Creates coordinates that default to the decorated cell.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            first_row_index: None,
            first_column_index: None,
            last_row_index: None,
            last_column_index: None,
            relative_first_row_index: None,
            relative_first_column_index: None,
            relative_last_row_index: None,
            relative_last_column_index: None,
        }
    }

    /// Sets the absolute first row. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn first_row_index(mut self, value: u32) -> Self {
        self.first_row_index = Some(value);
        self
    }

    /// Sets the absolute first column. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn first_column_index(mut self, value: u16) -> Self {
        self.first_column_index = Some(value);
        self
    }

    /// Sets the absolute last row. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn last_row_index(mut self, value: u32) -> Self {
        self.last_row_index = Some(value);
        self
    }

    /// Sets the absolute last column. Like Java, zero defers to the relative coordinate.
    #[must_use]
    pub const fn last_column_index(mut self, value: u16) -> Self {
        self.last_column_index = Some(value);
        self
    }

    /// Sets the first row relative to the decorated cell.
    #[must_use]
    pub const fn relative_first_row_index(mut self, value: i32) -> Self {
        self.relative_first_row_index = Some(value);
        self
    }

    /// Sets the first column relative to the decorated cell.
    #[must_use]
    pub const fn relative_first_column_index(mut self, value: i32) -> Self {
        self.relative_first_column_index = Some(value);
        self
    }

    /// Sets the last row relative to the decorated cell.
    #[must_use]
    pub const fn relative_last_row_index(mut self, value: i32) -> Self {
        self.relative_last_row_index = Some(value);
        self
    }

    /// Sets the last column relative to the decorated cell.
    #[must_use]
    pub const fn relative_last_column_index(mut self, value: i32) -> Self {
        self.relative_last_column_index = Some(value);
        self
    }

    /// Returns the absolute first row.
    #[must_use]
    pub const fn get_first_row_index(self) -> Option<u32> {
        self.first_row_index
    }

    /// Returns the absolute first column.
    #[must_use]
    pub const fn get_first_column_index(self) -> Option<u16> {
        self.first_column_index
    }

    /// Returns the absolute last row.
    #[must_use]
    pub const fn get_last_row_index(self) -> Option<u32> {
        self.last_row_index
    }

    /// Returns the absolute last column.
    #[must_use]
    pub const fn get_last_column_index(self) -> Option<u16> {
        self.last_column_index
    }

    /// Returns the relative first row.
    #[must_use]
    pub const fn get_relative_first_row_index(self) -> Option<i32> {
        self.relative_first_row_index
    }

    /// Returns the relative first column.
    #[must_use]
    pub const fn get_relative_first_column_index(self) -> Option<i32> {
        self.relative_first_column_index
    }

    /// Returns the relative last row.
    #[must_use]
    pub const fn get_relative_last_row_index(self) -> Option<i32> {
        self.relative_last_row_index
    }

    /// Returns the relative last column.
    #[must_use]
    pub const fn get_relative_last_column_index(self) -> Option<i32> {
        self.relative_last_column_index
    }
}

/// Java `ClientAnchorData.AnchorType` equivalent.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AnchorType {
    /// Move and resize with the anchor cells.
    #[default]
    MoveAndResize,
    /// POI's completeness-only mode; XLSX serializes it as a one-cell anchor.
    DontMoveDoResize,
    /// Move with cells without resizing.
    MoveDontResize,
    /// Do not move or resize with cells.
    DontMoveAndResize,
}

/// Client-anchor margins and movement behavior.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ClientAnchorData {
    coordinates: CoordinateData,
    top: Option<u32>,
    right: Option<u32>,
    bottom: Option<u32>,
    left: Option<u32>,
    anchor_type: Option<AnchorType>,
}

impl ClientAnchorData {
    /// Creates a default anchor for the decorated cell.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            coordinates: CoordinateData::new(),
            top: None,
            right: None,
            bottom: None,
            left: None,
            anchor_type: None,
        }
    }

    /// Sets its absolute and relative cell coordinates.
    #[must_use]
    pub const fn coordinates(mut self, value: CoordinateData) -> Self {
        self.coordinates = value;
        self
    }

    /// Sets the top margin in pixels.
    #[must_use]
    pub const fn top(mut self, value: u32) -> Self {
        self.top = Some(value);
        self
    }

    /// Sets the right margin in pixels.
    #[must_use]
    pub const fn right(mut self, value: u32) -> Self {
        self.right = Some(value);
        self
    }

    /// Sets the bottom margin in pixels.
    #[must_use]
    pub const fn bottom(mut self, value: u32) -> Self {
        self.bottom = Some(value);
        self
    }

    /// Sets the left margin in pixels.
    #[must_use]
    pub const fn left(mut self, value: u32) -> Self {
        self.left = Some(value);
        self
    }

    /// Sets the object movement and resize behavior.
    #[must_use]
    pub const fn anchor_type(mut self, value: AnchorType) -> Self {
        self.anchor_type = Some(value);
        self
    }

    /// Returns the coordinates.
    #[must_use]
    pub const fn get_coordinates(self) -> CoordinateData {
        self.coordinates
    }

    /// Returns the top margin in pixels.
    #[must_use]
    pub const fn get_top(self) -> Option<u32> {
        self.top
    }

    /// Returns the right margin in pixels.
    #[must_use]
    pub const fn get_right(self) -> Option<u32> {
        self.right
    }

    /// Returns the bottom margin in pixels.
    #[must_use]
    pub const fn get_bottom(self) -> Option<u32> {
        self.bottom
    }

    /// Returns the left margin in pixels.
    #[must_use]
    pub const fn get_left(self) -> Option<u32> {
        self.left
    }

    /// Returns the movement and resize behavior.
    #[must_use]
    pub const fn get_anchor_type(self) -> Option<AnchorType> {
        self.anchor_type
    }
}

/// Java `ImageData.ImageType` equivalent metadata.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageType {
    /// Extended Windows metafile.
    Emf,
    /// Windows metafile.
    Wmf,
    /// Macintosh PICT.
    Pict,
    /// JPEG.
    Jpeg,
    /// PNG.
    Png,
    /// Device-independent bitmap.
    Dib,
}

/// One Java-compatible image and its client anchor.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ImageData {
    image: Vec<u8>,
    image_type: Option<ImageType>,
    anchor: ClientAnchorData,
}

impl ImageData {
    /// Creates image data from encoded bytes.
    #[must_use]
    pub fn new(image: impl Into<Vec<u8>>) -> Self {
        Self {
            image: image.into(),
            image_type: None,
            anchor: ClientAnchorData::new(),
        }
    }

    /// Sets optional Java image-type metadata.
    #[must_use]
    pub const fn image_type(mut self, value: ImageType) -> Self {
        self.image_type = Some(value);
        self
    }

    /// Sets the client anchor.
    #[must_use]
    pub const fn anchor(mut self, value: ClientAnchorData) -> Self {
        self.anchor = value;
        self
    }

    /// Returns the encoded image bytes.
    #[must_use]
    pub fn image(&self) -> &[u8] {
        &self.image
    }

    /// Returns the optional image-type metadata.
    #[must_use]
    pub const fn get_image_type(&self) -> Option<ImageType> {
        self.image_type
    }

    /// Returns the client anchor.
    #[must_use]
    pub const fn get_anchor(&self) -> ClientAnchorData {
        self.anchor
    }
}

/// Runtime font metadata equivalent to Java `WriteFont`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct WriteFont {
    font_name: Option<String>,
    font_height_in_points: Option<f64>,
    italic: Option<bool>,
    strikeout: Option<bool>,
    color: Option<ExcelColor>,
    type_offset: Option<ExcelFontScript>,
    underline: Option<ExcelUnderline>,
    charset: Option<u8>,
    bold: Option<bool>,
}

impl WriteFont {
    /// Creates font metadata with every property unspecified.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            font_name: None,
            font_height_in_points: None,
            italic: None,
            strikeout: None,
            color: None,
            type_offset: None,
            underline: None,
            charset: None,
            bold: None,
        }
    }

    /// Sets the font family name.
    #[must_use]
    pub fn font_name(mut self, value: impl Into<String>) -> Self {
        self.font_name = Some(value.into());
        self
    }

    /// Sets the font size in points.
    #[must_use]
    pub const fn font_height_in_points(mut self, value: f64) -> Self {
        self.font_height_in_points = Some(value);
        self
    }

    /// Sets italic rendering.
    #[must_use]
    pub const fn italic(mut self, value: bool) -> Self {
        self.italic = Some(value);
        self
    }

    /// Sets strike-through rendering.
    #[must_use]
    pub const fn strikeout(mut self, value: bool) -> Self {
        self.strikeout = Some(value);
        self
    }

    /// Sets an indexed or RGB font color.
    #[must_use]
    pub const fn color(mut self, value: ExcelColor) -> Self {
        self.color = Some(value);
        self
    }

    /// Sets superscript or subscript rendering.
    #[must_use]
    pub const fn type_offset(mut self, value: ExcelFontScript) -> Self {
        self.type_offset = Some(value);
        self
    }

    /// Sets underline rendering.
    #[must_use]
    pub const fn underline(mut self, value: ExcelUnderline) -> Self {
        self.underline = Some(value);
        self
    }

    /// Sets the font character set.
    #[must_use]
    pub const fn charset(mut self, value: u8) -> Self {
        self.charset = Some(value);
        self
    }

    /// Sets bold rendering.
    #[must_use]
    pub const fn bold(mut self, value: bool) -> Self {
        self.bold = Some(value);
        self
    }

    /// Returns the optional font family name.
    #[must_use]
    pub fn get_font_name(&self) -> Option<&str> {
        self.font_name.as_deref()
    }

    /// Returns the optional font size.
    #[must_use]
    pub const fn get_font_height_in_points(&self) -> Option<f64> {
        self.font_height_in_points
    }

    /// Returns the optional italic flag.
    #[must_use]
    pub const fn get_italic(&self) -> Option<bool> {
        self.italic
    }

    /// Returns the optional strike-through flag.
    #[must_use]
    pub const fn get_strikeout(&self) -> Option<bool> {
        self.strikeout
    }

    /// Returns the optional font color.
    #[must_use]
    pub const fn get_color(&self) -> Option<ExcelColor> {
        self.color
    }

    /// Returns the optional superscript/subscript mode.
    #[must_use]
    pub const fn get_type_offset(&self) -> Option<ExcelFontScript> {
        self.type_offset
    }

    /// Returns the optional underline mode.
    #[must_use]
    pub const fn get_underline(&self) -> Option<ExcelUnderline> {
        self.underline
    }

    /// Returns the optional character set.
    #[must_use]
    pub const fn get_charset(&self) -> Option<u8> {
        self.charset
    }

    /// Returns the optional bold flag.
    #[must_use]
    pub const fn get_bold(&self) -> Option<bool> {
        self.bold
    }
}

/// One Java `RichTextStringData.IntervalFont` range using UTF-16 indices.
#[derive(Debug, Clone, PartialEq)]
pub struct IntervalFont {
    start_index: usize,
    end_index: usize,
    write_font: WriteFont,
}

impl IntervalFont {
    /// Creates a half-open font range `[start_index, end_index)`.
    #[must_use]
    pub const fn new(start_index: usize, end_index: usize, write_font: WriteFont) -> Self {
        Self {
            start_index,
            end_index,
            write_font,
        }
    }

    /// Returns the inclusive UTF-16 start index.
    #[must_use]
    pub const fn start_index(&self) -> usize {
        self.start_index
    }

    /// Returns the exclusive UTF-16 end index.
    #[must_use]
    pub const fn end_index(&self) -> usize {
        self.end_index
    }

    /// Returns the interval font.
    #[must_use]
    pub const fn write_font(&self) -> &WriteFont {
        &self.write_font
    }
}

/// Java `RichTextStringData` equivalent.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RichTextStringData {
    text_string: String,
    write_font: Option<WriteFont>,
    interval_font_list: Vec<IntervalFont>,
}

impl RichTextStringData {
    /// Creates rich-text metadata for a string.
    #[must_use]
    pub fn new(text_string: impl Into<String>) -> Self {
        Self {
            text_string: text_string.into(),
            write_font: None,
            interval_font_list: Vec::new(),
        }
    }

    /// Applies a font to the entire string.
    #[must_use]
    pub fn apply_font(mut self, write_font: WriteFont) -> Self {
        self.write_font = Some(write_font);
        self
    }

    /// Applies a font to a half-open UTF-16 character range.
    #[must_use]
    pub fn apply_font_range(
        mut self,
        start_index: usize,
        end_index: usize,
        write_font: WriteFont,
    ) -> Self {
        self.interval_font_list
            .push(IntervalFont::new(start_index, end_index, write_font));
        self
    }

    /// Replaces all interval font entries.
    #[must_use]
    pub fn interval_font_list(mut self, value: impl IntoIterator<Item = IntervalFont>) -> Self {
        self.interval_font_list = value.into_iter().collect();
        self
    }

    /// Returns the underlying text.
    #[must_use]
    pub fn text_string(&self) -> &str {
        &self.text_string
    }

    /// Returns the optional whole-string font.
    #[must_use]
    pub const fn write_font(&self) -> Option<&WriteFont> {
        self.write_font.as_ref()
    }

    /// Returns interval fonts in application order.
    #[must_use]
    pub fn interval_fonts(&self) -> &[IntervalFont] {
        &self.interval_font_list
    }
}

impl IntoExcelCell for RichTextStringData {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::RichText(self.clone()))
    }
}

impl FromExcelCell for RichTextStringData {
    fn from_excel_cell(cell: Option<&CellValue>, _context: &ConvertContext) -> Result<Self> {
        Ok(Self::new(cell.map_or_else(String::new, CellValue::as_text)))
    }
}

/// Java `WriteCellData` subset that preserves a scalar plus `imageDataList`.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellData {
    value: CellValue,
    image_data_list: Vec<ImageData>,
}

impl WriteCellData {
    /// Creates decorated cell data from a scalar value.
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

    /// Appends one image entry.
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

    /// Returns the scalar cell value.
    #[must_use]
    pub const fn value(&self) -> &CellValue {
        &self.value
    }

    /// Returns all image entries in insertion order.
    #[must_use]
    pub fn images(&self) -> &[ImageData] {
        &self.image_data_list
    }
}

impl IntoExcelCell for WriteCellData {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Images {
            value: Box::new(self.value.clone()),
            images: self.image_data_list.clone(),
        })
    }
}

impl FromExcelCell for WriteCellData {
    fn from_excel_cell(cell: Option<&CellValue>, _context: &ConvertContext) -> Result<Self> {
        Ok(Self::new(cell.cloned().unwrap_or(CellValue::Empty)))
    }
}

/// Logical Excel cell type used as the read-converter dispatch key.
///
/// The core variants mirror Java `EasyExcel`'s `CellDataTypeEnum`. `Formula` and
/// `Image` preserve Rust's richer backend-neutral values when a caller supplies
/// them directly.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CellDataType {
    /// Shared or inline string.
    String,
    /// Direct inline string.
    DirectString,
    /// Numeric value.
    Number,
    /// Boolean value.
    Boolean,
    /// Empty or physically absent cell.
    Empty,
    /// Excel error value.
    Error,
    /// Date or date-time value.
    Date,
    /// Rich text string.
    RichTextString,
    /// Formula expression supplied as a write value.
    Formula,
    /// Encoded image bytes.
    Image,
}

/// Formula metadata associated with a cached cell value while reading.
///
/// This mirrors Java `EasyExcel`'s `FormulaData`: the formula is kept separately
/// so typed conversion continues to consume Excel's cached result.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FormulaData {
    formula_value: String,
}

impl FormulaData {
    /// Creates formula metadata from the expression stored in the workbook.
    #[must_use]
    pub fn new(formula_value: impl Into<String>) -> Self {
        Self {
            formula_value: formula_value.into(),
        }
    }

    /// Returns the formula expression without adding a leading equals sign.
    #[must_use]
    pub fn formula_value(&self) -> &str {
        &self.formula_value
    }
}

/// Value mode used when reading rows without a declared Rust model.
///
/// This mirrors Java `EasyExcel`'s `ReadDefaultReturnEnum` while [`DynamicValue`]
/// keeps the runtime alternatives type-safe.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ReadDefaultReturn {
    /// Convert every present cell to the text a user sees in the workbook.
    #[default]
    String,
    /// Preserve the backend-neutral scalar type of each cell.
    ActualData,
    /// Return the scalar together with its raw value, location, and formula.
    ReadCellData,
}

/// Java-compatible no-model cell metadata.
#[derive(Debug, Clone, PartialEq)]
pub struct ReadCellData {
    row_index: u32,
    column_index: usize,
    raw_value: CellValue,
    data: CellValue,
    display_value: String,
    formula: Option<FormulaData>,
}

impl ReadCellData {
    fn new(
        row_index: u32,
        column_index: usize,
        raw_value: CellValue,
        data: CellValue,
        display_value: String,
        formula: Option<FormulaData>,
    ) -> Self {
        Self {
            row_index,
            column_index,
            raw_value,
            data,
            display_value,
            formula,
        }
    }

    /// Returns the physical zero-based row index.
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the physical zero-based column index.
    #[must_use]
    pub const fn column_index(&self) -> usize {
        self.column_index
    }

    /// Returns the original backend-neutral cell value.
    #[must_use]
    pub const fn raw_value(&self) -> &CellValue {
        &self.raw_value
    }

    /// Returns the Java `ACTUAL_DATA`-equivalent value.
    #[must_use]
    pub const fn data(&self) -> &CellValue {
        &self.data
    }

    /// Returns the Java-compatible formatted display text.
    #[must_use]
    pub fn display_value(&self) -> &str {
        &self.display_value
    }

    /// Returns formula metadata when the cell contains a formula.
    #[must_use]
    pub const fn formula(&self) -> Option<&FormulaData> {
        self.formula.as_ref()
    }
}

/// A type-safe value in a Java-compatible no-model row.
#[derive(Debug, Clone, PartialEq)]
pub enum DynamicValue {
    /// A missing column inserted to preserve physical indexes or head width.
    Null,
    /// Text returned by Java's default `STRING` mode.
    String(String),
    /// Scalar returned by Java's `ACTUAL_DATA` mode.
    ActualData(CellValue),
    /// Metadata returned by Java's `READ_CELL_DATA` mode.
    ReadCellData(ReadCellData),
}

/// A no-model row keyed by zero-based physical column index.
///
/// Use this when Java code would read `Map<Integer, String>`,
/// `Map<Integer, Object>`, or `Map<Integer, ReadCellData<?>>`.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct DynamicRow(BTreeMap<usize, DynamicValue>);

impl DynamicRow {
    /// Creates a dynamic row from indexed values.
    #[must_use]
    pub const fn new(values: BTreeMap<usize, DynamicValue>) -> Self {
        Self(values)
    }

    /// Returns all indexed values in physical column order.
    #[must_use]
    pub const fn values(&self) -> &BTreeMap<usize, DynamicValue> {
        &self.0
    }

    /// Returns a value by zero-based physical column index.
    #[must_use]
    pub fn get(&self, column_index: usize) -> Option<&DynamicValue> {
        self.0.get(&column_index)
    }

    /// Consumes the row and returns its ordered values.
    #[must_use]
    pub fn into_values(self) -> BTreeMap<usize, DynamicValue> {
        self.0
    }
}

/// Extra worksheet information selectable during a read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CellExtraType {
    /// A cell comment/note.
    Comment,
    /// A cell or range hyperlink.
    Hyperlink,
    /// A merged-cell range.
    Merge,
}

/// Extra worksheet information equivalent to Java `EasyExcel`'s `CellExtra`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellExtra {
    extra_type: CellExtraType,
    text: Option<String>,
    first_row_index: u32,
    last_row_index: u32,
    first_column_index: usize,
    last_column_index: usize,
}

impl CellExtra {
    /// Creates a cell or range extra event.
    #[must_use]
    pub fn new(
        extra_type: CellExtraType,
        text: Option<String>,
        first_row_index: u32,
        last_row_index: u32,
        first_column_index: usize,
        last_column_index: usize,
    ) -> Self {
        Self {
            extra_type,
            text,
            first_row_index,
            last_row_index,
            first_column_index,
            last_column_index,
        }
    }

    /// Returns the extra-data kind.
    #[must_use]
    pub const fn extra_type(&self) -> CellExtraType {
        self.extra_type
    }

    /// Returns comment text or hyperlink target; merge events have no text.
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    /// Returns the first zero-based row index.
    #[must_use]
    pub const fn first_row_index(&self) -> u32 {
        self.first_row_index
    }

    /// Returns the last zero-based row index.
    #[must_use]
    pub const fn last_row_index(&self) -> u32 {
        self.last_row_index
    }

    /// Returns the first zero-based column index.
    #[must_use]
    pub const fn first_column_index(&self) -> usize {
        self.first_column_index
    }

    /// Returns the last zero-based column index.
    #[must_use]
    pub const fn last_column_index(&self) -> usize {
        self.last_column_index
    }
}

/// Horizontal alignment used by annotation-driven cell styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelHorizontalAlignment {
    /// Excel's type-dependent default.
    General,
    /// Left aligned.
    Left,
    /// Centered.
    Center,
    /// Right aligned.
    Right,
    /// Repeats content across the cell.
    Fill,
    /// Justified.
    Justify,
    /// Centered across adjacent cells.
    CenterAcross,
    /// Distributed across the cell.
    Distributed,
}

/// Vertical alignment used by annotation-driven cell styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelVerticalAlignment {
    /// Top aligned.
    Top,
    /// Vertically centered.
    Center,
    /// Bottom aligned.
    Bottom,
    /// Vertically justified.
    Justify,
    /// Vertically distributed.
    Distributed,
}

/// Border line style used by annotation-driven cell styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelBorderStyle {
    /// No border.
    None,
    /// Thin solid line.
    Thin,
    /// Medium solid line.
    Medium,
    /// Dashed line.
    Dashed,
    /// Dotted line.
    Dotted,
    /// Thick solid line.
    Thick,
    /// Double line.
    Double,
    /// Hairline border.
    Hair,
    /// Medium dashed line.
    MediumDashed,
    /// Dash-dot line.
    DashDot,
    /// Medium dash-dot line.
    MediumDashDot,
    /// Dash-dot-dot line.
    DashDotDot,
    /// Medium dash-dot-dot line.
    MediumDashDotDot,
    /// Slanted dash-dot line.
    SlantDashDot,
}

/// Fill pattern used by annotation-driven cell styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelFillPattern {
    /// No fill.
    None,
    /// Solid foreground fill.
    Solid,
    /// 50% gray pattern.
    MediumGray,
    /// 75% gray pattern.
    DarkGray,
    /// 25% gray pattern.
    LightGray,
    /// Dark horizontal stripes.
    DarkHorizontal,
    /// Dark vertical stripes.
    DarkVertical,
    /// Dark downward diagonal stripes.
    DarkDown,
    /// Dark upward diagonal stripes.
    DarkUp,
    /// Dark grid.
    DarkGrid,
    /// Dark trellis.
    DarkTrellis,
    /// Light horizontal stripes.
    LightHorizontal,
    /// Light vertical stripes.
    LightVertical,
    /// Light downward diagonal stripes.
    LightDown,
    /// Light upward diagonal stripes.
    LightUp,
    /// Light grid.
    LightGrid,
    /// Light trellis.
    LightTrellis,
    /// 12.5% gray pattern.
    Gray125,
    /// 6.25% gray pattern.
    Gray0625,
}

/// Font underline style used by annotation-driven font styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelUnderline {
    /// No underline.
    None,
    /// Single underline.
    Single,
    /// Double underline.
    Double,
    /// Single accounting underline.
    SingleAccounting,
    /// Double accounting underline.
    DoubleAccounting,
}

/// Font script position used by annotation-driven font styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelFontScript {
    /// Normal baseline text.
    None,
    /// Superscript text.
    Superscript,
    /// Subscript text.
    Subscript,
}

/// A color supplied by Java-compatible palette index or by explicit RGB value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelColor {
    /// A Java EasyExcel/Apache POI indexed palette color.
    Indexed(u8),
    /// A backend-neutral RGB color in `0xRRGGBB` form.
    Rgb(u32),
}

impl ExcelColor {
    /// Interprets Java palette indexes `0..=64` as indexed colors and larger values as RGB.
    #[must_use]
    pub const fn java_or_rgb(value: u32) -> Self {
        if value <= 64 {
            Self::Indexed(value.to_le_bytes()[0])
        } else {
            Self::Rgb(value)
        }
    }
}

/// A Java built-in number-format index or a custom Excel format string.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelDataFormat {
    /// A Java EasyExcel/Apache POI built-in format index.
    Builtin(u8),
    /// A custom Excel number-format string.
    Custom(&'static str),
}

/// Cell-style properties generated from `HeadStyle` or `ContentStyle` equivalents.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ExcelCellStyle {
    /// Whether the cell is hidden when the sheet is protected.
    pub hidden: Option<bool>,
    /// Whether the cell is locked when the sheet is protected.
    pub locked: Option<bool>,
    /// Whether Excel treats the value as explicitly quoted text.
    pub quote_prefix: Option<bool>,
    /// Horizontal alignment.
    pub horizontal_alignment: Option<ExcelHorizontalAlignment>,
    /// Whether text wraps within the cell.
    pub wrapped: Option<bool>,
    /// Vertical alignment.
    pub vertical_alignment: Option<ExcelVerticalAlignment>,
    /// Text rotation in degrees.
    pub rotation: Option<i16>,
    /// Text indentation level.
    pub indent: Option<u8>,
    /// Left border style.
    pub border_left: Option<ExcelBorderStyle>,
    /// Right border style.
    pub border_right: Option<ExcelBorderStyle>,
    /// Top border style.
    pub border_top: Option<ExcelBorderStyle>,
    /// Bottom border style.
    pub border_bottom: Option<ExcelBorderStyle>,
    /// Left border indexed or RGB color.
    pub left_border_color: Option<ExcelColor>,
    /// Right border indexed or RGB color.
    pub right_border_color: Option<ExcelColor>,
    /// Top border indexed or RGB color.
    pub top_border_color: Option<ExcelColor>,
    /// Bottom border indexed or RGB color.
    pub bottom_border_color: Option<ExcelColor>,
    /// Fill pattern.
    pub fill_pattern: Option<ExcelFillPattern>,
    /// Fill background indexed or RGB color.
    pub fill_background_color: Option<ExcelColor>,
    /// Fill foreground indexed or RGB color.
    pub fill_foreground_color: Option<ExcelColor>,
    /// Whether text shrinks to fit the cell.
    pub shrink_to_fit: Option<bool>,
    /// Built-in or custom Excel number format.
    pub data_format: Option<ExcelDataFormat>,
}

impl ExcelCellStyle {
    /// Creates an annotation style with every property unspecified.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            hidden: None,
            locked: None,
            quote_prefix: None,
            horizontal_alignment: None,
            wrapped: None,
            vertical_alignment: None,
            rotation: None,
            indent: None,
            border_left: None,
            border_right: None,
            border_top: None,
            border_bottom: None,
            left_border_color: None,
            right_border_color: None,
            top_border_color: None,
            bottom_border_color: None,
            fill_pattern: None,
            fill_background_color: None,
            fill_foreground_color: None,
            shrink_to_fit: None,
            data_format: None,
        }
    }
}

/// Font properties generated from `HeadFontStyle` or `ContentFontStyle` equivalents.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ExcelFontStyle {
    /// Font family name.
    pub font_name: Option<&'static str>,
    /// Font size in points.
    pub font_height_in_points: Option<f64>,
    /// Italic rendering.
    pub italic: Option<bool>,
    /// Strike-through rendering.
    pub strikeout: Option<bool>,
    /// Font indexed or RGB color.
    pub color: Option<ExcelColor>,
    /// Superscript or subscript positioning.
    pub type_offset: Option<ExcelFontScript>,
    /// Underline rendering.
    pub underline: Option<ExcelUnderline>,
    /// Font character set.
    pub charset: Option<u8>,
    /// Bold rendering.
    pub bold: Option<bool>,
}

impl ExcelFontStyle {
    /// Creates an annotation font style with every property unspecified.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            font_name: None,
            font_height_in_points: None,
            italic: None,
            strikeout: None,
            color: None,
            type_offset: None,
            underline: None,
            charset: None,
            bold: None,
        }
    }
}

/// Static metadata for one Rust struct field and Excel column.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ExcelColumn {
    /// Rust field name.
    pub field: &'static str,
    /// Excel header name.
    pub name: &'static str,
    /// Explicit zero-based column index.
    pub index: Option<usize>,
    /// Relative ordering when no explicit index is configured.
    pub order: i32,
    /// Optional date or number format.
    pub format: Option<&'static str>,
    /// Optional annotation-driven column width in Excel character units.
    pub column_width: Option<u16>,
    /// Field-level header cell style.
    pub head_style: Option<ExcelCellStyle>,
    /// Field-level content cell style.
    pub content_style: Option<ExcelCellStyle>,
    /// Field-level header font style.
    pub head_font_style: Option<ExcelFontStyle>,
    /// Field-level content font style.
    pub content_font_style: Option<ExcelFontStyle>,
}

impl ExcelColumn {
    /// Creates static column metadata.
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
        }
    }

    /// Creates static column metadata with annotation-driven write dimensions.
    #[must_use]
    pub const fn with_column_width(mut self, width: u16) -> Self {
        self.column_width = Some(width);
        self
    }

    /// Adds a field-level header cell style.
    #[must_use]
    pub const fn with_head_style(mut self, style: ExcelCellStyle) -> Self {
        self.head_style = Some(style);
        self
    }

    /// Adds a field-level content cell style.
    #[must_use]
    pub const fn with_content_style(mut self, style: ExcelCellStyle) -> Self {
        self.content_style = Some(style);
        self
    }

    /// Adds a field-level header font style.
    #[must_use]
    pub const fn with_head_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.head_font_style = Some(style);
        self
    }

    /// Adds a field-level content font style.
    #[must_use]
    pub const fn with_content_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.content_font_style = Some(style);
        self
    }
}

/// Type-level dimensions derived from Java-style write annotations.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ExcelWriteMetadata {
    /// Default width for columns without a field-level override.
    pub column_width: Option<u16>,
    /// Height for every generated header row.
    pub head_row_height: Option<u16>,
    /// Height for every generated content row.
    pub content_row_height: Option<u16>,
    /// Type-level header cell style.
    pub head_style: Option<ExcelCellStyle>,
    /// Type-level content cell style.
    pub content_style: Option<ExcelCellStyle>,
    /// Type-level header font style.
    pub head_font_style: Option<ExcelFontStyle>,
    /// Type-level content font style.
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

    /// Sets the type-level default column width.
    #[must_use]
    pub const fn column_width(mut self, width: u16) -> Self {
        self.column_width = Some(width);
        self
    }

    /// Sets the generated header-row height.
    #[must_use]
    pub const fn head_row_height(mut self, height: u16) -> Self {
        self.head_row_height = Some(height);
        self
    }

    /// Sets the generated content-row height.
    #[must_use]
    pub const fn content_row_height(mut self, height: u16) -> Self {
        self.content_row_height = Some(height);
        self
    }

    /// Adds a type-level header cell style.
    #[must_use]
    pub const fn head_style(mut self, style: ExcelCellStyle) -> Self {
        self.head_style = Some(style);
        self
    }

    /// Adds a type-level content cell style.
    #[must_use]
    pub const fn content_style(mut self, style: ExcelCellStyle) -> Self {
        self.content_style = Some(style);
        self
    }

    /// Adds a type-level header font style.
    #[must_use]
    pub const fn head_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.head_font_style = Some(style);
        self
    }

    /// Adds a type-level content font style.
    #[must_use]
    pub const fn content_font_style(mut self, style: ExcelFontStyle) -> Self {
        self.content_font_style = Some(style);
        self
    }
}

/// A physical row plus resolved header positions.
#[derive(Debug, Clone)]
pub struct RowData {
    sheet_name: String,
    row_index: u32,
    cells: Vec<CellValue>,
    headers: Arc<HashMap<String, usize>>,
    formulas: HashMap<usize, FormulaData>,
    display_values: HashMap<usize, String>,
    decimal_values: HashMap<usize, BigDecimal>,
    present_columns: HashSet<usize>,
    read_default_return: ReadDefaultReturn,
}

impl RowData {
    /// Creates row data.
    #[must_use]
    pub fn new(
        sheet_name: impl Into<String>,
        row_index: u32,
        cells: Vec<CellValue>,
        headers: Arc<HashMap<String, usize>>,
    ) -> Self {
        let present_columns = (0..cells.len()).collect();
        Self {
            sheet_name: sheet_name.into(),
            row_index,
            cells,
            headers,
            formulas: HashMap::new(),
            display_values: HashMap::new(),
            decimal_values: HashMap::new(),
            present_columns,
            read_default_return: ReadDefaultReturn::default(),
        }
    }

    /// Attaches formula metadata indexed by zero-based physical column.
    #[must_use]
    pub fn with_formulas(mut self, formulas: HashMap<usize, FormulaData>) -> Self {
        self.formulas = formulas;
        self
    }

    /// Attaches Java-compatible formatted display text by physical column index.
    #[must_use]
    pub fn with_display_values(mut self, display_values: HashMap<usize, String>) -> Self {
        self.display_values = display_values;
        self
    }

    /// Attaches exact OOXML decimal values by physical column index.
    #[must_use]
    pub fn with_decimal_values(mut self, decimal_values: HashMap<usize, BigDecimal>) -> Self {
        self.decimal_values = decimal_values;
        self
    }

    /// Attaches the physical columns that were explicitly present in the source.
    #[must_use]
    pub fn with_present_columns(mut self, present_columns: HashSet<usize>) -> Self {
        self.present_columns = present_columns;
        self
    }

    /// Selects the Java-compatible no-model return mode.
    #[must_use]
    pub const fn with_read_default_return(mut self, mode: ReadDefaultReturn) -> Self {
        self.read_default_return = mode;
        self
    }

    /// Returns the physical zero-based row index.
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the sheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Resolves a cell using Java `EasyExcel`'s index-before-name priority.
    #[must_use]
    pub fn cell(&self, column: &ExcelColumn) -> Option<&CellValue> {
        let index = column
            .index
            .or_else(|| self.headers.get(column.name).copied())?;
        self.cells.get(index)
    }

    /// Resolves formula metadata using the same index-before-name priority as [`Self::cell`].
    #[must_use]
    pub fn formula(&self, column: &ExcelColumn) -> Option<&FormulaData> {
        let index = column
            .index
            .or_else(|| self.headers.get(column.name).copied())?;
        self.formulas.get(&index)
    }

    fn dynamic_width(&self) -> usize {
        let head_width = self
            .headers
            .values()
            .copied()
            .max()
            .map_or(0, |index| index.saturating_add(1));
        self.cells.len().max(head_width)
    }

    fn dynamic_cell(&self, column_index: usize) -> DynamicValue {
        if !self.present_columns.contains(&column_index) {
            return DynamicValue::Null;
        }
        let raw_value = self
            .cells
            .get(column_index)
            .cloned()
            .unwrap_or(CellValue::Empty);
        let raw_value = if matches!(raw_value, CellValue::Int(_) | CellValue::Float(_)) {
            self.decimal_values
                .get(&column_index)
                .cloned()
                .map_or(raw_value, CellValue::Decimal)
        } else {
            raw_value
        };
        let data = actual_cell_value(&raw_value);
        let display_value = self
            .display_values
            .get(&column_index)
            .cloned()
            .unwrap_or_else(|| raw_value.as_text());
        match self.read_default_return {
            ReadDefaultReturn::String => DynamicValue::String(display_value),
            ReadDefaultReturn::ActualData => DynamicValue::ActualData(data),
            ReadDefaultReturn::ReadCellData => DynamicValue::ReadCellData(ReadCellData::new(
                self.row_index,
                column_index,
                raw_value,
                data,
                display_value,
                self.formulas.get(&column_index).cloned(),
            )),
        }
    }

    /// Creates a conversion context for a column.
    #[must_use]
    pub fn convert_context(&self, column: &ExcelColumn) -> ConvertContext {
        let column_index = column
            .index
            .or_else(|| self.headers.get(column.name).copied());
        ConvertContext {
            sheet_name: self.sheet_name.clone(),
            row_index: self.row_index,
            column_index,
            field: column.field,
            format: column.format,
        }
    }
}

/// Location and formatting information supplied to cell converters.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConvertContext {
    /// Sheet name.
    pub sheet_name: String,
    /// Zero-based row index.
    pub row_index: u32,
    /// Zero-based column index when it can be resolved.
    pub column_index: Option<usize>,
    /// Rust field name.
    pub field: &'static str,
    /// Optional format string.
    pub format: Option<&'static str>,
}

impl ConvertContext {
    fn invalid(&self, value: &CellValue, target: &'static str) -> ExcelError {
        ExcelError::Data {
            sheet: self.sheet_name.clone(),
            row: self.row_index,
            column: self.column_index,
            field: self.field,
            value: value.as_text(),
            message: format!("cannot convert cell to {target}"),
        }
    }
}

/// Converts a backend-neutral cell into a Rust value.
pub trait FromExcelCell: Sized {
    /// Performs the conversion.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Data`] when the cell cannot be represented by `Self`.
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self>;
}

/// Converts a Rust value into a backend-neutral cell.
pub trait IntoExcelCell {
    /// Performs the conversion.
    ///
    /// # Errors
    ///
    /// Returns an error when the Rust value cannot be represented by an Excel cell.
    fn to_excel_cell(&self, context: &ConvertContext) -> Result<CellValue>;
}

/// Context supplied to a custom cell-to-Rust converter.
#[derive(Debug, Clone, Copy)]
pub struct ReadConverterContext<'a> {
    cell: Option<&'a CellValue>,
    formula: Option<&'a FormulaData>,
    column: &'a ExcelColumn,
    context: &'a ConvertContext,
}

impl<'a> ReadConverterContext<'a> {
    /// Creates a read conversion context.
    #[must_use]
    pub const fn new(
        cell: Option<&'a CellValue>,
        column: &'a ExcelColumn,
        context: &'a ConvertContext,
    ) -> Self {
        Self {
            cell,
            formula: None,
            column,
            context,
        }
    }

    /// Creates a read conversion context with optional formula metadata.
    #[must_use]
    pub const fn with_formula(
        cell: Option<&'a CellValue>,
        formula: Option<&'a FormulaData>,
        column: &'a ExcelColumn,
        context: &'a ConvertContext,
    ) -> Self {
        Self {
            cell,
            formula,
            column,
            context,
        }
    }

    /// Returns the source cell, or `None` when it is physically absent.
    #[must_use]
    pub const fn cell(&self) -> Option<&'a CellValue> {
        self.cell
    }

    /// Returns formula metadata when the source cell contains a formula.
    #[must_use]
    pub const fn formula(&self) -> Option<&'a FormulaData> {
        self.formula
    }

    /// Returns the field's static column metadata.
    #[must_use]
    pub const fn column(&self) -> &'a ExcelColumn {
        self.column
    }

    /// Returns the resolved row, column, field, and format information.
    #[must_use]
    pub const fn convert_context(&self) -> &'a ConvertContext {
        self.context
    }
}

/// Context supplied to a custom Rust-to-cell converter.
#[derive(Debug, Clone, Copy)]
pub struct WriteConverterContext<'a, T> {
    value: &'a T,
    column: &'a ExcelColumn,
    context: &'a ConvertContext,
}

impl<'a, T> WriteConverterContext<'a, T> {
    /// Creates a write conversion context.
    #[must_use]
    pub const fn new(value: &'a T, column: &'a ExcelColumn, context: &'a ConvertContext) -> Self {
        Self {
            value,
            column,
            context,
        }
    }

    /// Returns the Rust field value.
    #[must_use]
    pub const fn value(&self) -> &'a T {
        self.value
    }

    /// Returns the field's static column metadata.
    #[must_use]
    pub const fn column(&self) -> &'a ExcelColumn {
        self.column
    }

    /// Returns the target row, column, field, and format information.
    #[must_use]
    pub const fn convert_context(&self) -> &'a ConvertContext {
        self.context
    }
}

/// Custom bidirectional converter selected by `#[excel(converter = Type)]`.
#[allow(clippy::missing_errors_doc)]
pub trait Converter<T> {
    /// Returns the source cell type supported when this converter is registered globally.
    ///
    /// Java `EasyExcel` requires global read converters to expose this key. A
    /// string default keeps field-only converters concise while matching the
    /// most common custom converter contract.
    fn support_excel_type(&self) -> CellDataType {
        CellDataType::String
    }

    /// Converts an Excel cell into a Rust field value.
    fn convert_to_rust_data(&self, _context: &ReadConverterContext<'_>) -> Result<T> {
        Err(ExcelError::Unsupported(
            "custom converter does not support reading".to_owned(),
        ))
    }

    /// Converts a Rust field value into an Excel cell.
    fn convert_to_excel_data(&self, _context: &WriteConverterContext<'_, T>) -> Result<CellValue> {
        Err(ExcelError::Unsupported(
            "custom converter does not support writing".to_owned(),
        ))
    }
}

/// Java `StringImageConverter` equivalent for fields containing an image file path.
///
/// Use it with `#[excel(converter = StringImageConverter)]`. The file is read
/// during row conversion; missing or unreadable files return an I/O error.
#[derive(Debug, Clone, Copy, Default)]
pub struct StringImageConverter;

impl Converter<String> for StringImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<CellValue> {
        std::fs::read(context.value())
            .map(CellValue::Image)
            .map_err(Into::into)
    }
}

/// Java `InputStreamImageConverter` equivalent for a stateful Rust [`Read`] source.
///
/// Conversion consumes the bytes remaining in the reader and deliberately does
/// not close or replace it, matching Java `EasyExcel`'s ownership contract for a
/// caller-supplied `InputStream`.
#[derive(Debug)]
pub struct ImageInputStream<R> {
    reader: RefCell<R>,
}

impl<R> ImageInputStream<R> {
    /// Wraps a reader whose remaining bytes represent one image.
    #[must_use]
    pub const fn new(reader: R) -> Self {
        Self {
            reader: RefCell::new(reader),
        }
    }

    /// Returns the wrapped reader, preserving its position after conversion.
    #[must_use]
    pub fn into_inner(self) -> R {
        self.reader.into_inner()
    }
}

impl<R> From<R> for ImageInputStream<R> {
    fn from(reader: R) -> Self {
        Self::new(reader)
    }
}

impl<R: Read> IntoExcelCell for ImageInputStream<R> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        read_image_bytes(&mut *self.reader.borrow_mut()).map(CellValue::Image)
    }
}

fn read_image_bytes(reader: &mut dyn Read) -> Result<Vec<u8>> {
    let mut bytes = Vec::new();
    reader.read_to_end(&mut bytes)?;
    Ok(bytes)
}

/// Java `InputStreamImageConverter` equivalent for annotation-selected stream fields.
#[derive(Debug, Clone, Copy, Default)]
pub struct InputStreamImageConverter;

impl<R: Read> Converter<ImageInputStream<R>> for InputStreamImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, ImageInputStream<R>>,
    ) -> Result<CellValue> {
        context.value().to_excel_cell(context.convert_context())
    }
}

/// Java `UrlImageConverter` equivalent with Java's default timeout values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UrlImageConverter {
    connect_timeout: Duration,
    read_timeout: Duration,
}

impl UrlImageConverter {
    /// Java `EasyExcel`'s default URL connection timeout.
    pub const DEFAULT_CONNECT_TIMEOUT: Duration = Duration::from_secs(1);
    /// Java `EasyExcel`'s default URL response-read timeout.
    pub const DEFAULT_READ_TIMEOUT: Duration = Duration::from_secs(5);

    /// Creates a converter with explicit connection and response-read timeouts.
    #[must_use]
    pub const fn new(connect_timeout: Duration, read_timeout: Duration) -> Self {
        Self {
            connect_timeout,
            read_timeout,
        }
    }

    /// Returns the configured connection timeout.
    #[must_use]
    pub fn connect_timeout(self) -> Duration {
        self.connect_timeout
    }

    /// Returns the configured response-read timeout.
    #[must_use]
    pub fn read_timeout(self) -> Duration {
        self.read_timeout
    }

    fn download(self, value: &Url) -> Result<Vec<u8>> {
        let agent: ureq::Agent = ureq::Agent::config_builder()
            .timeout_connect(Some(self.connect_timeout))
            .timeout_recv_body(Some(self.read_timeout))
            .build()
            .into();
        let mut response = agent.get(value.as_str()).call().map_err(url_image_error)?;
        let mut bytes = Vec::new();
        response
            .body_mut()
            .as_reader()
            .read_to_end(&mut bytes)
            .map_err(url_image_error)?;
        Ok(bytes)
    }
}

impl Default for UrlImageConverter {
    fn default() -> Self {
        Self::new(Self::DEFAULT_CONNECT_TIMEOUT, Self::DEFAULT_READ_TIMEOUT)
    }
}

impl Converter<Url> for UrlImageConverter {
    fn convert_to_excel_data(&self, context: &WriteConverterContext<'_, Url>) -> Result<CellValue> {
        self.download(context.value()).map(CellValue::Image)
    }
}

impl IntoExcelCell for Url {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        UrlImageConverter::default()
            .download(self)
            .map(CellValue::Image)
    }
}

fn url_image_error(error: impl Display) -> ExcelError {
    ExcelError::Io(std::io::Error::other(error.to_string()))
}

trait ErasedConverter: Send + Sync {
    fn target_type_id(&self) -> TypeId;
    fn target_type_name(&self) -> &'static str;
    fn support_excel_type(&self) -> CellDataType;
    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<Box<dyn Any>>;
    fn convert_to_excel_data(
        &self,
        value: &dyn Any,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<CellValue>;
}

struct TypedConverter<T, C> {
    converter: C,
    marker: std::marker::PhantomData<fn() -> T>,
}

impl<T, C> ErasedConverter for TypedConverter<T, C>
where
    T: 'static,
    C: Converter<T> + Send + Sync,
{
    fn target_type_id(&self) -> TypeId {
        TypeId::of::<T>()
    }

    fn target_type_name(&self) -> &'static str {
        type_name::<T>()
    }

    fn support_excel_type(&self) -> CellDataType {
        self.converter.support_excel_type()
    }

    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<Box<dyn Any>> {
        self.converter
            .convert_to_rust_data(context)
            .map(|value| Box::new(value) as Box<dyn Any>)
    }

    fn convert_to_excel_data(
        &self,
        value: &dyn Any,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<CellValue> {
        let value = value.downcast_ref::<T>().ok_or_else(|| {
            ExcelError::Format(format!(
                "registered converter expected Rust type {}",
                type_name::<T>()
            ))
        })?;
        self.converter
            .convert_to_excel_data(&WriteConverterContext::new(value, column, context))
    }
}

/// Runtime converter registry populated by Java-style `registerConverter` builders.
///
/// Registrations are searched from newest to oldest. Read selection uses the
/// pair `(Rust target type, Excel cell type)` while write selection uses only
/// the Rust type, matching Java `EasyExcel`'s holder initialization rules.
#[derive(Clone, Default)]
pub struct ConverterRegistry {
    converters: Vec<Arc<dyn ErasedConverter>>,
}

impl ConverterRegistry {
    /// Registers a converter for `T`, overriding an earlier converter with the same key.
    pub fn register<T, C>(&mut self, converter: C)
    where
        T: 'static,
        C: Converter<T> + Send + Sync + 'static,
    {
        self.converters.push(Arc::new(TypedConverter::<T, C> {
            converter,
            marker: std::marker::PhantomData,
        }));
    }

    /// Returns a registry where `overrides` take precedence over this registry.
    #[must_use]
    pub fn merged_with(&self, overrides: &Self) -> Self {
        let mut converters = self.converters.clone();
        converters.extend(overrides.converters.iter().cloned());
        Self { converters }
    }

    /// Converts a cell through the newest matching global converter.
    ///
    /// `None` means no global converter matched and the caller should use its
    /// built-in conversion implementation.
    ///
    /// # Errors
    ///
    /// Returns the registered converter's error or a type-contract error.
    pub fn convert_to_rust_data<T>(&self, context: &ReadConverterContext<'_>) -> Result<Option<T>>
    where
        T: 'static,
    {
        let data_type = context
            .cell()
            .map_or(CellDataType::Empty, CellValue::data_type);
        let Some(converter) = self.converters.iter().rev().find(|converter| {
            converter.target_type_id() == TypeId::of::<T>()
                && converter.support_excel_type() == data_type
        }) else {
            return Ok(None);
        };
        converter
            .convert_to_rust_data(context)?
            .downcast::<T>()
            .map(|value| Some(*value))
            .map_err(|_| {
                ExcelError::Format(format!(
                    "registered converter returned a value other than {}",
                    type_name::<T>()
                ))
            })
    }

    /// Converts a Rust value through the newest matching global converter.
    ///
    /// `None` means no global converter matched and the caller should use its
    /// built-in conversion implementation.
    ///
    /// # Errors
    ///
    /// Returns the registered converter's conversion error.
    pub fn convert_to_excel_data<T>(
        &self,
        value: &T,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<Option<CellValue>>
    where
        T: 'static,
    {
        let Some(converter) = self
            .converters
            .iter()
            .rev()
            .find(|converter| converter.target_type_id() == TypeId::of::<T>())
        else {
            return Ok(None);
        };
        converter
            .convert_to_excel_data(value, column, context)
            .map(Some)
    }

    /// Returns whether no custom converter has been registered.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.converters.is_empty()
    }
}

impl std::fmt::Debug for ConverterRegistry {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_list()
            .entries(
                self.converters.iter().map(|converter| {
                    (converter.target_type_name(), converter.support_excel_type())
                }),
            )
            .finish()
    }
}

impl PartialEq for ConverterRegistry {
    fn eq(&self, other: &Self) -> bool {
        self.converters.len() == other.converters.len()
            && self
                .converters
                .iter()
                .zip(&other.converters)
                .all(|(left, right)| {
                    left.target_type_id() == right.target_type_id()
                        && left.support_excel_type() == right.support_excel_type()
                })
    }
}

impl Eq for ConverterRegistry {}

/// Workbook-level write lifecycle context.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteWorkbookContext {
    path: PathBuf,
}

impl WriteWorkbookContext {
    /// Creates a workbook context for an output path.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        Self { path: path.into() }
    }

    /// Returns the output path.
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }
}

/// Worksheet-level write lifecycle context.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheetContext {
    sheet_name: String,
}

impl WriteSheetContext {
    /// Creates a worksheet context.
    #[must_use]
    pub fn new(sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
        }
    }

    /// Returns the worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }
}

/// Row-level write lifecycle context.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteRowContext {
    /// Worksheet name.
    pub sheet_name: String,
    /// Physical zero-based row index.
    pub row_index: u32,
    /// Whether this is a header row.
    pub is_head: bool,
}

/// Mutable cell-level write lifecycle context.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellContext {
    /// Worksheet name.
    pub sheet_name: String,
    /// Physical zero-based row index.
    pub row_index: u32,
    /// Physical zero-based column index.
    pub column_index: u16,
    /// Rust field name, when backed by a typed column.
    pub field: Option<&'static str>,
    /// Whether this is a header cell.
    pub is_head: bool,
    /// Value that will be written. A handler may replace it.
    pub value: CellValue,
    /// A handler may set this to suppress the physical cell.
    pub skip: bool,
}

/// Intercepts the workbook, worksheet, row, and cell write lifecycle.
///
/// Any callback may return an error to stop the write immediately.
#[allow(clippy::missing_errors_doc)]
pub trait WriteHandler {
    /// Lower orders execute first.
    fn order(&self) -> i32 {
        0
    }

    /// Called before the workbook is created.
    fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Called after the workbook is saved.
    fn after_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Called after the worksheet is configured and before rows are written.
    fn before_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Called after all worksheet rows are written.
    fn after_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Called before a row is written.
    fn before_row(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Called after a row is written.
    fn after_row(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Called before a cell is written and may transform or skip it.
    fn before_cell(&mut self, _context: &mut WriteCellContext) -> Result<()> {
        Ok(())
    }

    /// Called after a cell has been processed.
    fn after_cell(&mut self, _context: &WriteCellContext) -> Result<()> {
        Ok(())
    }
}

impl FromExcelCell for String {
    fn from_excel_cell(value: Option<&CellValue>, _context: &ConvertContext) -> Result<Self> {
        Ok(value.map_or_else(String::new, CellValue::as_text))
    }
}

impl IntoExcelCell for String {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::String(self.clone()))
    }
}

impl IntoExcelCell for &str {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::String((*self).to_owned()))
    }
}

impl FromExcelCell for bool {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        match value.unwrap_or(&CellValue::Empty) {
            CellValue::Bool(value) => Ok(*value),
            CellValue::Int(value) => Ok(*value != 0),
            CellValue::Float(value) => Ok(*value != 0.0),
            CellValue::Decimal(value) => Ok(value != &BigDecimal::from(0)),
            CellValue::String(value) if value.eq_ignore_ascii_case("true") || value == "1" => {
                Ok(true)
            }
            CellValue::String(value) if value.eq_ignore_ascii_case("false") || value == "0" => {
                Ok(false)
            }
            other => Err(context.invalid(other, "bool")),
        }
    }
}

impl IntoExcelCell for bool {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Bool(*self))
    }
}

macro_rules! integer_conversion {
    ($($ty:ty),+ $(,)?) => {
        $(
            impl FromExcelCell for $ty {
                fn from_excel_cell(
                    value: Option<&CellValue>,
                    context: &ConvertContext,
                ) -> Result<Self> {
                    parse_integer(value, context, stringify!($ty))
                }
            }

            impl IntoExcelCell for $ty {
                fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
                    Ok(integer_to_cell(*self))
                }
            }
        )+
    };
}

integer_conversion!(i8, i16, i32, i64, isize, u8, u16, u32, u64, usize);

impl FromExcelCell for BigInt {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        let cell = value.unwrap_or(&CellValue::Empty);
        match cell {
            CellValue::Bool(value) => Ok(Self::from(u8::from(*value))),
            CellValue::Int(value) => Ok(Self::from(*value)),
            CellValue::Float(value) => BigDecimal::from_str(&value.to_string())
                .map(|value| decimal_to_big_int(&value))
                .map_err(|_| context.invalid(cell, "BigInt")),
            CellValue::Decimal(value) => Ok(decimal_to_big_int(value)),
            CellValue::String(value) => BigDecimal::from_str(value)
                .map(|value| decimal_to_big_int(&value))
                .map_err(|_| context.invalid(cell, "BigInt")),
            other => Err(context.invalid(other, "BigInt")),
        }
    }
}

impl IntoExcelCell for BigInt {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(self
            .to_i64()
            .map_or_else(|| CellValue::String(self.to_string()), CellValue::Int))
    }
}

fn decimal_to_big_int(value: &BigDecimal) -> BigInt {
    value.with_scale(0).into_bigint_and_exponent().0
}

fn parse_integer<T>(
    value: Option<&CellValue>,
    context: &ConvertContext,
    target: &'static str,
) -> Result<T>
where
    T: FromStr,
{
    let value = value.unwrap_or(&CellValue::Empty);
    let text = match value {
        CellValue::Int(inner) => inner.to_string(),
        CellValue::Float(inner) if inner.fract() == 0.0 => inner.to_string(),
        CellValue::Decimal(inner) if inner == &inner.with_scale(0) => inner.to_string(),
        CellValue::String(inner) => inner.clone(),
        other => return Err(context.invalid(other, target)),
    };
    text.parse::<T>()
        .map_err(|_| context.invalid(value, target))
}

fn integer_to_cell<T>(value: T) -> CellValue
where
    T: TryInto<i64> + Display + Copy,
{
    value
        .try_into()
        .map_or_else(|_| CellValue::String(value.to_string()), CellValue::Int)
}

macro_rules! float_conversion {
    ($($ty:ty),+ $(,)?) => {
        $(
            impl FromExcelCell for $ty {
                fn from_excel_cell(
                    value: Option<&CellValue>,
                    context: &ConvertContext,
                ) -> Result<Self> {
                    parse_float(value, context, stringify!($ty))
                }
            }

            impl IntoExcelCell for $ty {
                fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
                    Ok(CellValue::Float(f64::from(*self)))
                }
            }
        )+
    };
}

float_conversion!(f32, f64);

fn parse_float<T>(
    value: Option<&CellValue>,
    context: &ConvertContext,
    target: &'static str,
) -> Result<T>
where
    T: FromStr,
{
    let value = value.unwrap_or(&CellValue::Empty);
    let text = match value {
        CellValue::Int(inner) => inner.to_string(),
        CellValue::Float(inner) => inner.to_string(),
        CellValue::Decimal(inner) => inner.to_string(),
        CellValue::String(inner) => inner.clone(),
        other => return Err(context.invalid(other, target)),
    };
    text.parse::<T>()
        .map_err(|_| context.invalid(value, target))
}

impl FromExcelCell for BigDecimal {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::Decimal(inner) => Ok(inner.clone()),
            CellValue::Int(inner) => Ok(Self::from(*inner)),
            CellValue::Float(inner) => {
                Self::from_str(&inner.to_string()).map_err(|_| context.invalid(value, "BigDecimal"))
            }
            CellValue::String(inner) => {
                Self::from_str(inner).map_err(|_| context.invalid(value, "BigDecimal"))
            }
            other => Err(context.invalid(other, "BigDecimal")),
        }
    }
}

impl IntoExcelCell for BigDecimal {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Decimal(self.clone()))
    }
}

impl FromExcelCell for NaiveDate {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::Date(value) => Ok(*value),
            CellValue::DateTime(value) => Ok(value.date()),
            CellValue::String(inner) => {
                NaiveDate::parse_from_str(inner, context.format.unwrap_or("%Y-%m-%d"))
                    .map_err(|_| context.invalid(value, "NaiveDate"))
            }
            other => Err(context.invalid(other, "NaiveDate")),
        }
    }
}

impl IntoExcelCell for NaiveDate {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Date(*self))
    }
}

impl FromExcelCell for NaiveDateTime {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::DateTime(value) => Ok(*value),
            CellValue::Date(value) => Ok(value.and_hms_opt(0, 0, 0).expect("midnight is valid")),
            CellValue::String(inner) => {
                NaiveDateTime::parse_from_str(inner, context.format.unwrap_or("%Y-%m-%d %H:%M:%S"))
                    .map_err(|_| context.invalid(value, "NaiveDateTime"))
            }
            other => Err(context.invalid(other, "NaiveDateTime")),
        }
    }
}

impl IntoExcelCell for NaiveDateTime {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::DateTime(*self))
    }
}

impl FromExcelCell for Vec<u8> {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::Image(bytes) => Ok(bytes.clone()),
            other => Err(context.invalid(other, "Vec<u8>")),
        }
    }
}

impl IntoExcelCell for Vec<u8> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Image(self.clone()))
    }
}

impl FromExcelCell for Box<[u8]> {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        Vec::<u8>::from_excel_cell(value, context).map(Vec::into_boxed_slice)
    }
}

impl IntoExcelCell for Box<[u8]> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Image(self.to_vec()))
    }
}

impl<const N: usize> FromExcelCell for [u8; N] {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        Vec::<u8>::from_excel_cell(value, context)?
            .try_into()
            .map_err(|_| context.invalid(value.unwrap_or(&CellValue::Empty), "[u8; N]"))
    }
}

impl<const N: usize> IntoExcelCell for [u8; N] {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        Ok(CellValue::Image(self.to_vec()))
    }
}

impl FromExcelCell for PathBuf {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        String::from_excel_cell(value, context).map(Self::from)
    }
}

impl IntoExcelCell for PathBuf {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue> {
        std::fs::read(self)
            .map(CellValue::Image)
            .map_err(Into::into)
    }
}

impl<T: FromExcelCell> FromExcelCell for Option<T> {
    fn from_excel_cell(value: Option<&CellValue>, context: &ConvertContext) -> Result<Self> {
        if value.is_none_or(CellValue::is_empty) {
            Ok(None)
        } else {
            T::from_excel_cell(value, context).map(Some)
        }
    }
}

impl<T: IntoExcelCell> IntoExcelCell for Option<T> {
    fn to_excel_cell(&self, context: &ConvertContext) -> Result<CellValue> {
        self.as_ref()
            .map_or(Ok(CellValue::Empty), |value| value.to_excel_cell(context))
    }
}

/// Compile-time mapping implemented by `#[derive(ExcelRow)]`.
pub trait ExcelRow: Sized {
    /// Returns static column metadata.
    fn schema() -> &'static [ExcelColumn];

    /// Returns annotation-driven dimensions used by writers.
    #[must_use]
    fn write_metadata() -> &'static ExcelWriteMetadata {
        const METADATA: ExcelWriteMetadata = ExcelWriteMetadata::new();
        &METADATA
    }

    /// Converts a physical row into the user type.
    ///
    /// # Errors
    ///
    /// Returns a location-aware conversion error when any field cannot be decoded.
    fn from_row(row: &RowData) -> Result<Self>;

    /// Converts a physical row using Java-style globally registered converters.
    ///
    /// Implementations that do not expose typed fields can retain the default.
    ///
    /// # Errors
    ///
    /// Returns a location-aware conversion error.
    fn from_row_with_converters(row: &RowData, _converters: &ConverterRegistry) -> Result<Self> {
        Self::from_row(row)
    }

    /// Converts the user type into schema-ordered cells.
    ///
    /// # Errors
    ///
    /// Returns an error when any field cannot be encoded as an Excel cell.
    fn to_row(&self) -> Result<Vec<CellValue>>;

    /// Converts the value using Java-style globally registered converters.
    ///
    /// Implementations that do not expose typed fields can retain the default.
    ///
    /// # Errors
    ///
    /// Returns an error when a field cannot be represented as an Excel cell.
    fn to_row_with_converters(&self, _converters: &ConverterRegistry) -> Result<Vec<CellValue>> {
        self.to_row()
    }
}

impl ExcelRow for DynamicRow {
    fn schema() -> &'static [ExcelColumn] {
        &[]
    }

    fn from_row(row: &RowData) -> Result<Self> {
        Ok(Self(
            (0..row.dynamic_width())
                .map(|index| (index, row.dynamic_cell(index)))
                .collect(),
        ))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        let Some(last_index) = self.0.last_key_value().map(|(index, _)| *index) else {
            return Ok(Vec::new());
        };
        let row_length = last_index
            .checked_add(1)
            .ok_or_else(|| ExcelError::Format("dynamic column index exceeds usize".to_owned()))?;
        let mut row = vec![CellValue::Empty; row_length];
        for (index, value) in &self.0 {
            row[*index] = match value {
                DynamicValue::Null => CellValue::Empty,
                DynamicValue::String(value) => CellValue::String(value.clone()),
                DynamicValue::ActualData(value) => value.clone(),
                DynamicValue::ReadCellData(value) => value.data().clone(),
            };
        }
        Ok(row)
    }
}

fn actual_cell_value(value: &CellValue) -> CellValue {
    match value {
        CellValue::Empty => CellValue::String(String::new()),
        CellValue::Error(value) => CellValue::String(value.clone()),
        value => value.clone(),
    }
}

/// Type-safe shared value equivalent to Java `EasyExcel`'s reader `customObject`.
#[derive(Clone)]
pub struct CustomReadObject(Arc<dyn Any + Send + Sync>);

impl CustomReadObject {
    /// Wraps a value for propagation to every read callback context.
    #[must_use]
    pub fn new<T>(value: T) -> Self
    where
        T: Any + Send + Sync,
    {
        Self(Arc::new(value))
    }

    /// Returns the value when its concrete type matches `T`.
    #[must_use]
    pub fn downcast_ref<T: Any>(&self) -> Option<&T> {
        self.0.downcast_ref()
    }
}

impl std::fmt::Debug for CustomReadObject {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CustomReadObject")
            .finish_non_exhaustive()
    }
}

impl PartialEq for CustomReadObject {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl Eq for CustomReadObject {}

/// Read callback context equivalent to Java `AnalysisContext`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnalysisContext {
    sheet_name: String,
    sheet_no: usize,
    row_index: u32,
    batch_index: usize,
    custom_object: Option<CustomReadObject>,
}

impl AnalysisContext {
    /// Creates a context.
    #[must_use]
    pub fn new(sheet_name: impl Into<String>, sheet_no: usize, row_index: u32) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            sheet_no,
            row_index,
            batch_index: 0,
            custom_object: None,
        }
    }

    /// Returns the sheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Returns the zero-based sheet index.
    #[must_use]
    pub const fn sheet_no(&self) -> usize {
        self.sheet_no
    }

    /// Returns the zero-based physical row index.
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the zero-based callback batch index.
    #[must_use]
    pub const fn batch_index(&self) -> usize {
        self.batch_index
    }

    /// Returns the configured custom read object, if any.
    #[must_use]
    pub const fn custom_object(&self) -> Option<&CustomReadObject> {
        self.custom_object.as_ref()
    }

    /// Returns the custom read object when its concrete type matches `T`.
    #[must_use]
    pub fn custom<T: Any>(&self) -> Option<&T> {
        self.custom_object.as_ref()?.downcast_ref()
    }

    /// Returns a context carrying the supplied custom read object.
    #[must_use]
    pub fn with_custom_object(mut self, custom_object: Option<CustomReadObject>) -> Self {
        self.custom_object = custom_object;
        self
    }

    /// Returns a copy with a different batch index.
    #[must_use]
    pub fn with_batch_index(&self, batch_index: usize) -> Self {
        let mut context = self.clone();
        context.batch_index = batch_index;
        context
    }
}

/// Action selected by a listener after a row error.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ErrorAction {
    /// Continue with the next row.
    Continue,
    /// Skip the failed row and continue.
    SkipRow,
    /// Stop and return the error.
    Stop,
}

/// Event listener equivalent to Java `EasyExcel`'s `ReadListener`.
pub trait ReadListener<T> {
    /// Called when row conversion or processing fails.
    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        ErrorAction::Stop
    }

    /// Called for a resolved header row.
    ///
    /// # Errors
    ///
    /// Returns an error to stop the read operation.
    fn invoke_head(
        &mut self,
        _head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        Ok(())
    }

    /// Called once for every successfully converted row.
    ///
    /// # Errors
    ///
    /// Returns an error to stop the read operation.
    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()>;

    /// Called when enabled comment, hyperlink, or merge metadata is encountered.
    ///
    /// # Errors
    ///
    /// Returns an error to route through [`Self::on_exception`].
    fn extra(&mut self, _extra: &CellExtra, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }

    /// Called after a sheet has been analysed.
    ///
    /// # Errors
    ///
    /// Returns an error when final listener work fails.
    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }

    /// Allows a listener to stop before the next row.
    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        true
    }
}

impl<T, L: ReadListener<T> + ?Sized> ReadListener<T> for Box<L> {
    fn on_exception(&mut self, error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        (**self).on_exception(error, context)
    }

    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        (**self).invoke_head(head, context)
    }

    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        (**self).invoke(data, context)
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        (**self).extra(extra, context)
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        (**self).do_after_all_analysed(context)
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        (**self).has_next(context)
    }
}

/// A listener that buffers rows and invokes a callback page by page.
type PageCallback<T> = dyn FnMut(Vec<T>, &AnalysisContext) -> Result<()>;

/// A listener that buffers rows and invokes a callback page by page.
pub struct PageReadListener<T> {
    batch_size: usize,
    batch_index: usize,
    rows: Vec<T>,
    callback: Box<PageCallback<T>>,
}

impl<T> PageReadListener<T> {
    /// Creates a paged listener. A zero size is normalized to one row.
    #[must_use]
    pub fn new(
        batch_size: usize,
        callback: impl FnMut(Vec<T>, &AnalysisContext) -> Result<()> + 'static,
    ) -> Self {
        let batch_size = batch_size.max(1);
        Self {
            batch_size,
            batch_index: 0,
            rows: Vec::with_capacity(batch_size),
            callback: Box::new(callback),
        }
    }

    fn flush(&mut self, context: &AnalysisContext) -> Result<()> {
        if self.rows.is_empty() {
            return Ok(());
        }
        let rows = std::mem::replace(&mut self.rows, Vec::with_capacity(self.batch_size));
        let context = context.with_batch_index(self.batch_index);
        complete_page(&mut self.batch_index, (self.callback)(rows, &context))
    }
}

impl<T> ReadListener<T> for PageReadListener<T> {
    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        if self.rows.len() >= self.batch_size {
            return self.flush(context);
        }
        Ok(())
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.flush(context)
    }
}

fn complete_page(batch_index: &mut usize, result: Result<()>) -> Result<()> {
    result.map(|()| {
        *batch_index += 1;
    })
}

/// All public easyexcel errors with row and column diagnostics where applicable.
#[derive(Debug, Error)]
pub enum ExcelError {
    /// A cell-to-field conversion error.
    #[error(
        "sheet={sheet}, row={row}, column={column:?}, field={field}, value={value:?}: {message}"
    )]
    Data {
        /// Sheet name.
        sheet: String,
        /// Zero-based row index.
        row: u32,
        /// Zero-based column index.
        column: Option<usize>,
        /// Rust field name.
        field: &'static str,
        /// Original cell text.
        value: String,
        /// Human-readable failure reason.
        message: String,
    },
    /// A requested worksheet does not exist.
    #[error("worksheet not found: {0}")]
    SheetNotFound(String),
    /// The workbook or OOXML package is invalid.
    #[error("excel format error: {0}")]
    Format(String),
    /// The requested operation is not supported by the selected engine.
    #[error("unsupported operation: {0}")]
    Unsupported(String),
    /// An I/O operation failed.
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

#[cfg(test)]
mod tests;
