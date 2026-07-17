//! Core data model and extension points for `easyexcel-rs`.

use std::collections::HashMap;
use std::fmt::Display;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::Arc;

use chrono::{NaiveDate, NaiveDateTime};
use thiserror::Error;

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
            Self::Date(value) => value.format("%Y-%m-%d").to_string(),
            Self::DateTime(value) => value.format("%Y-%m-%d %H:%M:%S").to_string(),
            Self::Hyperlink { text, .. } => text.clone(),
            Self::Comment { value, .. } => value.as_text(),
        }
    }
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
        Self {
            sheet_name: sheet_name.into(),
            row_index,
            cells,
            headers,
            formulas: HashMap::new(),
        }
    }

    /// Attaches formula metadata indexed by zero-based physical column.
    #[must_use]
    pub fn with_formulas(mut self, formulas: HashMap<usize, FormulaData>) -> Self {
        self.formulas = formulas;
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
        CellValue::String(inner) => inner.clone(),
        other => return Err(context.invalid(other, target)),
    };
    text.parse::<T>()
        .map_err(|_| context.invalid(value, target))
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

    /// Converts the user type into schema-ordered cells.
    ///
    /// # Errors
    ///
    /// Returns an error when any field cannot be encoded as an Excel cell.
    fn to_row(&self) -> Result<Vec<CellValue>>;
}

/// Read callback context equivalent to Java `AnalysisContext`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnalysisContext {
    sheet_name: String,
    sheet_no: usize,
    row_index: u32,
    batch_index: usize,
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
