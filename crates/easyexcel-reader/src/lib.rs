//! XLSX, XLS, and CSV readers backed by Calamine and the Rust CSV engine.

use std::collections::HashMap;
use std::path::Path;
use std::sync::Arc;

use calamine::{Data, DataRef, Range, Reader, Xls, Xlsx, open_workbook};
use easyexcel_core::{
    AnalysisContext, CellValue, ErrorAction, ExcelError, ExcelRow, ReadListener, Result, RowData,
};

/// Selects a worksheet by index, name, or all sheets.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub enum SheetSelector {
    /// The first worksheet.
    #[default]
    First,
    /// A zero-based worksheet index.
    Index(usize),
    /// A worksheet name.
    Name(String),
    /// Every worksheet in workbook order.
    All,
}

/// Workbook read configuration shared by XLSX, XLS, and CSV engines.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReadOptions {
    /// Sheet selection.
    pub sheet: SheetSelector,
    /// Number of header rows. The final header row is used for name mapping.
    pub head_row_number: u32,
    /// Whether rows containing only empty cells are ignored.
    pub ignore_empty_row: bool,
}

impl Default for ReadOptions {
    fn default() -> Self {
        Self {
            sheet: SheetSelector::First,
            head_row_number: 1,
            ignore_empty_row: true,
        }
    }
}

/// Reads selected XLSX sheets and dispatches typed row events.
///
/// # Errors
///
/// Returns an I/O, workbook-format, sheet-selection, conversion, or listener error.
pub fn read_xlsx<T, L>(path: &Path, options: &ReadOptions, listener: &mut L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let mut workbook: Xlsx<_> = open_workbook(path).map_err(format_error)?;
    let names = selected_sheet_names(&workbook, &options.sheet)?;
    for (sheet_no, sheet_name) in names {
        let mut consumer = TypedRowConsumer::<T> { listener };
        read_sheet(&mut workbook, sheet_no, &sheet_name, options, &mut consumer)?;
    }
    Ok(())
}

/// Reads selected legacy XLS sheets through the typed listener lifecycle.
///
/// Calamine materializes each XLS worksheet before row dispatch because the
/// binary BIFF format does not expose the XLSX cell-stream API.
///
/// # Errors
///
/// Returns an I/O, workbook-format, sheet-selection, conversion, or listener error.
pub fn read_xls<T, L>(path: &Path, options: &ReadOptions, listener: &mut L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let mut workbook: Xls<_> = open_workbook(path).map_err(format_error)?;
    let sheets = select_xls_sheets(workbook.worksheets(), &options.sheet)?;
    for (sheet_no, sheet_name, range) in sheets {
        let mut consumer = TypedRowConsumer::<T> { listener };
        read_range(&range, sheet_no, &sheet_name, options, &mut consumer)?;
    }
    Ok(())
}

/// Reads a CSV file through the same typed listener lifecycle as XLSX.
///
/// CSV exposes one logical sheet. Indexes other than zero return `SheetNotFound`.
///
/// # Errors
///
/// Returns an I/O, CSV-format, sheet-selection, conversion, or listener error.
pub fn read_csv<T, L>(path: &Path, options: &ReadOptions, listener: &mut L) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let sheet_name = csv_sheet_name(&options.sheet)?;
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_path(path)
        .map_err(format_error)?;
    read_csv_records::<T, L>(&mut reader.records(), 0, &sheet_name, options, listener)
}

fn read_csv_records<T, L>(
    records: &mut dyn Iterator<Item = csv::Result<csv::StringRecord>>,
    start_row: usize,
    sheet_name: &str,
    options: &ReadOptions,
    listener: &mut L,
) -> Result<()>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    let mut headers = Arc::new(HashMap::new());
    let mut final_row = 0_u32;
    for (offset, record) in records.enumerate() {
        let row_index = start_row.saturating_add(offset);
        let row_index = csv_row_index(row_index)?;
        final_row = row_index;
        let cells = record
            .map_err(format_error)?
            .iter()
            .map(|value| CellValue::String(value.to_owned()))
            .collect();
        process_row::<T>(
            0,
            sheet_name,
            row_index,
            cells,
            options,
            &mut headers,
            listener,
        )?;
    }
    listener.do_after_all_analysed(&AnalysisContext::new(sheet_name, 0, final_row))
}

fn csv_row_index(row_index: usize) -> Result<u32> {
    u32::try_from(row_index).map_err(|_| ExcelError::Format("CSV row index exceeds u32".to_owned()))
}

fn csv_sheet_name(selector: &SheetSelector) -> Result<String> {
    match selector {
        SheetSelector::First | SheetSelector::Index(0) | SheetSelector::All => {
            Ok("Sheet1".to_owned())
        }
        SheetSelector::Name(name) => Ok(name.clone()),
        SheetSelector::Index(index) => Err(ExcelError::SheetNotFound(index.to_string())),
    }
}

trait RowConsumer {
    fn process(
        &mut self,
        sheet_no: usize,
        sheet_name: &str,
        row_index: u32,
        cells: Vec<CellValue>,
        options: &ReadOptions,
        headers: &mut Arc<HashMap<String, usize>>,
    ) -> Result<()>;

    fn after(&mut self, context: &AnalysisContext) -> Result<()>;
}

struct TypedRowConsumer<'a, T> {
    listener: &'a mut dyn ReadListener<T>,
}

impl<T: ExcelRow> RowConsumer for TypedRowConsumer<'_, T> {
    fn process(
        &mut self,
        sheet_no: usize,
        sheet_name: &str,
        row_index: u32,
        cells: Vec<CellValue>,
        options: &ReadOptions,
        headers: &mut Arc<HashMap<String, usize>>,
    ) -> Result<()> {
        process_row::<T>(
            sheet_no,
            sheet_name,
            row_index,
            cells,
            options,
            headers,
            self.listener,
        )
    }

    fn after(&mut self, context: &AnalysisContext) -> Result<()> {
        self.listener.do_after_all_analysed(context)
    }
}

fn selected_sheet_names<RS: std::io::Read + std::io::Seek>(
    workbook: &Xlsx<RS>,
    selector: &SheetSelector,
) -> Result<Vec<(usize, String)>> {
    select_sheet_names(workbook.sheet_names(), selector)
}

fn select_sheet_names(
    names: Vec<String>,
    selector: &SheetSelector,
) -> Result<Vec<(usize, String)>> {
    match selector {
        SheetSelector::First => names
            .first()
            .cloned()
            .map(|name| vec![(0, name)])
            .ok_or_else(|| ExcelError::SheetNotFound("0".to_owned())),
        SheetSelector::Index(index) => names
            .get(*index)
            .cloned()
            .map(|name| vec![(*index, name)])
            .ok_or_else(|| ExcelError::SheetNotFound(index.to_string())),
        SheetSelector::Name(name) => names
            .iter()
            .position(|candidate| candidate == name)
            .map(|index| vec![(index, name.clone())])
            .ok_or_else(|| ExcelError::SheetNotFound(name.clone())),
        SheetSelector::All => Ok(names.into_iter().enumerate().collect()),
    }
}

fn select_xls_sheets(
    sheets: Vec<(String, Range<Data>)>,
    selector: &SheetSelector,
) -> Result<Vec<(usize, String, Range<Data>)>> {
    match selector {
        SheetSelector::First => sheets
            .into_iter()
            .next()
            .map(|(name, range)| vec![(0, name, range)])
            .ok_or_else(|| ExcelError::SheetNotFound("0".to_owned())),
        SheetSelector::Index(index) => sheets
            .into_iter()
            .nth(*index)
            .map(|(name, range)| vec![(*index, name, range)])
            .ok_or_else(|| ExcelError::SheetNotFound(index.to_string())),
        SheetSelector::Name(name) => sheets
            .into_iter()
            .enumerate()
            .find(|(_, (candidate, _))| candidate == name)
            .map(|(index, (_, range))| vec![(index, name.clone(), range)])
            .ok_or_else(|| ExcelError::SheetNotFound(name.clone())),
        SheetSelector::All => Ok(sheets
            .into_iter()
            .enumerate()
            .map(|(index, (name, range))| (index, name, range))
            .collect()),
    }
}

fn read_sheet<RS>(
    workbook: &mut Xlsx<RS>,
    sheet_no: usize,
    sheet_name: &str,
    options: &ReadOptions,
    consumer: &mut dyn RowConsumer,
) -> Result<()>
where
    RS: std::io::Read + std::io::Seek,
{
    let mut reader = workbook
        .worksheet_cells_reader(sheet_name)
        .map_err(format_error)?;
    let mut current_index = None;
    let mut current_cells = Vec::new();
    let mut headers = Arc::new(HashMap::new());

    while let Some(cell) = reader.next_cell().map_err(format_error)? {
        let (row, column) = cell.get_position();
        if current_index.is_some_and(|current| current != row) {
            consumer.process(
                sheet_no,
                sheet_name,
                current_index.expect("row index exists"),
                std::mem::take(&mut current_cells),
                options,
                &mut headers,
            )?;
        }
        current_index = Some(row);
        let column = to_column_index(column)?;
        if current_cells.len() <= column {
            current_cells.resize(column + 1, CellValue::Empty);
        }
        current_cells[column] = from_calamine(cell.get_value());
    }

    if let Some(row) = current_index {
        consumer.process(
            sheet_no,
            sheet_name,
            row,
            current_cells,
            options,
            &mut headers,
        )?;
    }

    let final_row = current_index.unwrap_or_default();
    consumer.after(&AnalysisContext::new(sheet_name, sheet_no, final_row))
}

fn read_range(
    range: &Range<Data>,
    sheet_no: usize,
    sheet_name: &str,
    options: &ReadOptions,
    consumer: &mut dyn RowConsumer,
) -> Result<()> {
    let mut headers = Arc::new(HashMap::new());
    let Some((start_row, start_column)) = range.start() else {
        return consumer.after(&AnalysisContext::new(sheet_name, sheet_no, 0));
    };
    let start_column = to_column_index(start_column)?;
    let mut row_index = start_row;
    let mut final_row = start_row;
    for row in range.rows() {
        final_row = row_index;
        let mut cells = vec![CellValue::Empty; start_column];
        cells.extend(row.iter().map(from_data));
        consumer.process(
            sheet_no,
            sheet_name,
            row_index,
            cells,
            options,
            &mut headers,
        )?;
        row_index = row_index.saturating_add(1);
    }
    consumer.after(&AnalysisContext::new(sheet_name, sheet_no, final_row))
}

#[allow(clippy::too_many_arguments)]
fn process_row<T>(
    sheet_no: usize,
    sheet_name: &str,
    row_index: u32,
    cells: Vec<CellValue>,
    options: &ReadOptions,
    headers: &mut Arc<HashMap<String, usize>>,
    listener: &mut dyn ReadListener<T>,
) -> Result<()>
where
    T: ExcelRow,
{
    let context = AnalysisContext::new(sheet_name, sheet_no, row_index);
    if options.head_row_number > 0 && row_index + 1 == options.head_row_number {
        *headers = Arc::new(header_map(&cells));
        return listener.invoke_head(headers, &context);
    }
    if row_index < options.head_row_number
        || (options.ignore_empty_row && cells.iter().all(CellValue::is_empty))
        || !listener.has_next(&context)
    {
        return Ok(());
    }

    let row = RowData::new(sheet_name, row_index, cells, Arc::clone(headers));
    match T::from_row(&row) {
        Ok(data) => listener.invoke(data, &context),
        Err(error) => match listener.on_exception(&error, &context) {
            ErrorAction::Continue | ErrorAction::SkipRow => Ok(()),
            ErrorAction::Stop => Err(error),
        },
    }
}

fn header_map(cells: &[CellValue]) -> HashMap<String, usize> {
    cells
        .iter()
        .enumerate()
        .filter_map(|(index, value)| {
            let name = value.as_text();
            (!name.is_empty()).then_some((name, index))
        })
        .collect()
}

fn from_calamine(value: &DataRef<'_>) -> CellValue {
    match value {
        DataRef::Empty => CellValue::Empty,
        DataRef::String(value) | DataRef::DateTimeIso(value) | DataRef::DurationIso(value) => {
            CellValue::String(value.clone())
        }
        DataRef::SharedString(value) => CellValue::String((*value).to_owned()),
        DataRef::Bool(value) => CellValue::Bool(*value),
        DataRef::Int(value) => CellValue::Int(*value),
        DataRef::Float(value) => CellValue::Float(*value),
        DataRef::DateTime(value) => {
            if value.is_datetime() {
                value
                    .as_datetime()
                    .map_or(CellValue::Float(value.as_f64()), CellValue::DateTime)
            } else {
                CellValue::Float(value.as_f64())
            }
        }
        DataRef::Error(value) => CellValue::Error(format!("{value:?}")),
    }
}

fn from_data(value: &Data) -> CellValue {
    match value {
        Data::Empty => CellValue::Empty,
        Data::String(value) | Data::DateTimeIso(value) | Data::DurationIso(value) => {
            CellValue::String(value.clone())
        }
        Data::Bool(value) => CellValue::Bool(*value),
        Data::Int(value) => CellValue::Int(*value),
        Data::Float(value) => CellValue::Float(*value),
        Data::DateTime(value) => {
            if value.is_datetime() {
                value
                    .as_datetime()
                    .map_or(CellValue::Float(value.as_f64()), CellValue::DateTime)
            } else {
                CellValue::Float(value.as_f64())
            }
        }
        Data::Error(value) => CellValue::Error(format!("{value:?}")),
    }
}

fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

fn to_column_index(column: u32) -> Result<usize> {
    u16::try_from(column)
        .map(usize::from)
        .map_err(|_| ExcelError::Format("column index exceeds spreadsheet limit".to_owned()))
}

#[cfg(test)]
mod tests;
