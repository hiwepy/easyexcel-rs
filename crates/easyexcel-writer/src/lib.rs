//! XLSX writer backed by `rust_xlsxwriter`.

use std::path::Path;

use easyexcel_core::{CellValue, ExcelColumn, ExcelError, ExcelRow, Result};
use rust_xlsxwriter::{Format, Workbook, Worksheet};

/// XLSX write configuration.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(clippy::struct_excessive_bools)]
pub struct WriteOptions {
    /// Worksheet name.
    pub sheet_name: String,
    /// Whether to use a one-row constant-memory worksheet.
    pub constant_memory: bool,
    /// Whether column headers are written.
    pub need_head: bool,
    /// Whether header rows are frozen.
    pub freeze_head: bool,
    /// Explicit freeze pane position as `(row, column)`.
    pub freeze_panes: Option<(u32, u16)>,
    /// Physical column indexes to include.
    pub include_column_indexes: Option<Vec<usize>>,
    /// Rust field names to include.
    pub include_column_field_names: Option<Vec<String>>,
    /// Physical column indexes to exclude.
    pub exclude_column_indexes: Vec<usize>,
    /// Rust field names to exclude.
    pub exclude_column_field_names: Vec<String>,
    /// Whether included columns follow the order of the include list.
    pub order_by_include_column: bool,
}

impl Default for WriteOptions {
    fn default() -> Self {
        Self {
            sheet_name: "Sheet1".to_owned(),
            constant_memory: false,
            need_head: true,
            freeze_head: false,
            freeze_panes: None,
            include_column_indexes: None,
            include_column_field_names: None,
            exclude_column_indexes: Vec::new(),
            exclude_column_field_names: Vec::new(),
            order_by_include_column: false,
        }
    }
}

/// Writes typed rows to a new XLSX file.
///
/// # Errors
///
/// Returns a conversion, worksheet-configuration, XLSX-format, or I/O error.
pub fn write_xlsx<T, I>(path: &Path, options: &WriteOptions, rows: I) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let mut workbook = Workbook::new();
    let worksheet = if options.constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(&options.sheet_name)
        .map_err(format_error)?;
    let freeze_panes = options
        .freeze_panes
        .or_else(|| (options.freeze_head && options.need_head).then_some((1, 0)));
    if let Some((row, column)) = freeze_panes {
        worksheet
            .set_freeze_panes(row, column)
            .map_err(format_error)?;
    }

    let columns = selected_columns(T::schema(), options);
    let mut row_index = 0_u32;
    if options.need_head {
        write_headers(worksheet, &columns)?;
        row_index = 1;
    }
    for row in rows {
        let cells = row.to_row()?;
        write_data_row(worksheet, row_index, &columns, &cells)?;
        row_index += 1;
    }
    workbook.save(path).map_err(format_error)
}

fn ordered_columns(schema: &'static [ExcelColumn]) -> Vec<(usize, usize, &'static ExcelColumn)> {
    let mut columns = schema
        .iter()
        .enumerate()
        .map(|(schema_index, column)| {
            let physical_index = column.index.unwrap_or(schema_index);
            (physical_index, schema_index, column)
        })
        .collect::<Vec<_>>();
    columns.sort_by_key(|(physical_index, schema_index, column)| {
        (*physical_index, column.order, *schema_index)
    });
    columns
}

fn selected_columns(
    schema: &'static [ExcelColumn],
    options: &WriteOptions,
) -> Vec<(usize, usize, &'static ExcelColumn)> {
    let mut columns = ordered_columns(schema)
        .into_iter()
        .filter(|(physical_index, _, column)| {
            let included_by_index = options
                .include_column_indexes
                .as_ref()
                .is_some_and(|indexes| indexes.contains(physical_index));
            let included_by_name = options
                .include_column_field_names
                .as_ref()
                .is_some_and(|names| names.iter().any(|name| name == column.field));
            let has_includes = options.include_column_indexes.is_some()
                || options.include_column_field_names.is_some();
            let excluded = options.exclude_column_indexes.contains(physical_index)
                || options
                    .exclude_column_field_names
                    .iter()
                    .any(|name| name == column.field);
            (!has_includes || included_by_index || included_by_name) && !excluded
        })
        .collect::<Vec<_>>();

    if options.order_by_include_column {
        columns.sort_by_key(|(physical_index, _, column)| {
            options
                .include_column_indexes
                .as_ref()
                .and_then(|indexes| indexes.iter().position(|index| index == physical_index))
                .or_else(|| {
                    options
                        .include_column_field_names
                        .as_ref()
                        .and_then(|names| names.iter().position(|name| name == column.field))
                })
                .unwrap_or(usize::MAX)
        });
        for (output_index, (physical_index, _, _)) in columns.iter_mut().enumerate() {
            *physical_index = output_index;
        }
    }
    columns
}

fn write_headers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
) -> Result<()> {
    let format = Format::new().set_bold();
    for (physical_index, _, column) in columns {
        worksheet
            .write_string_with_format(0, to_column(*physical_index)?, column.name, &format)
            .map_err(format_error)?;
    }
    Ok(())
}

fn write_data_row(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
) -> Result<()> {
    for (physical_index, schema_index, metadata) in columns {
        let value = cells.get(*schema_index).unwrap_or(&CellValue::Empty);
        let column = to_column(*physical_index)?;
        match value {
            CellValue::Empty => {}
            CellValue::String(value) | CellValue::Error(value) => {
                worksheet
                    .write_string(row_index, column, value)
                    .map_err(format_error)?;
            }
            CellValue::Bool(value) => {
                worksheet
                    .write_boolean(row_index, column, *value)
                    .map_err(format_error)?;
            }
            CellValue::Int(value) => {
                write_integer(worksheet, row_index, column, *value)?;
            }
            CellValue::Float(value) => {
                worksheet
                    .write_number(row_index, column, *value)
                    .map_err(format_error)?;
            }
            CellValue::Date(value) => {
                let format =
                    Format::new().set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd"));
                worksheet
                    .write_datetime_with_format(row_index, column, *value, &format)
                    .map_err(format_error)?;
            }
            CellValue::DateTime(value) => {
                let format = Format::new()
                    .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd hh:mm:ss"));
                worksheet
                    .write_datetime_with_format(row_index, column, *value, &format)
                    .map_err(format_error)?;
            }
        }
    }
    Ok(())
}

fn write_integer(worksheet: &mut Worksheet, row: u32, column: u16, value: i64) -> Result<()> {
    const MAX_EXACT_EXCEL_INTEGER: u64 = 9_007_199_254_740_991;
    if value.unsigned_abs() <= MAX_EXACT_EXCEL_INTEGER {
        #[allow(clippy::cast_precision_loss)]
        let number = value as f64;
        worksheet
            .write_number(row, column, number)
            .map_err(format_error)?;
    } else {
        worksheet
            .write_string(row, column, value.to_string())
            .map_err(format_error)?;
    }
    Ok(())
}

fn excel_date_format(format: Option<&str>, default: &str) -> String {
    format
        .unwrap_or(default)
        .replace("%Y", "yyyy")
        .replace("%m", "mm")
        .replace("%d", "dd")
        .replace("%H", "hh")
        .replace("%M", "mm")
        .replace("%S", "ss")
}

fn to_column(index: usize) -> Result<u16> {
    u16::try_from(index)
        .map_err(|_| ExcelError::Format("column index exceeds XLSX limit".to_owned()))
}

fn format_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[cfg(test)]
mod tests;
