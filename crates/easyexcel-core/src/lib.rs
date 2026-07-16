//! Core data model and extension points for `easyexcel-rs`.

use std::collections::HashMap;
use std::fmt::Display;
use std::str::FromStr;
use std::sync::Arc;

use chrono::{NaiveDate, NaiveDateTime};
use thiserror::Error;

/// The result type used by all easyexcel crates.
pub type Result<T> = std::result::Result<T, ExcelError>;

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
            Self::Empty => String::new(),
            Self::String(value) | Self::Error(value) => value.clone(),
            Self::Bool(value) => value.to_string(),
            Self::Int(value) => value.to_string(),
            Self::Float(value) => value.to_string(),
            Self::Date(value) => value.format("%Y-%m-%d").to_string(),
            Self::DateTime(value) => value.format("%Y-%m-%d %H:%M:%S").to_string(),
        }
    }
}

/// Static metadata for one Rust struct field and Excel column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
        }
    }
}

/// A physical row plus resolved header positions.
#[derive(Debug, Clone)]
pub struct RowData {
    sheet_name: String,
    row_index: u32,
    cells: Vec<CellValue>,
    headers: Arc<HashMap<String, usize>>,
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
        }
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
        match (self.callback)(rows, &context) {
            Ok(()) => {
                self.batch_index += 1;
                Ok(())
            }
            Err(error) => Err(error),
        }
    }
}

impl<T> ReadListener<T> for PageReadListener<T> {
    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        if self.rows.len() >= self.batch_size {
            match self.flush(context) {
                Ok(()) => {}
                Err(error) => return Err(error),
            }
        }
        Ok(())
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.flush(context)
    }
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
