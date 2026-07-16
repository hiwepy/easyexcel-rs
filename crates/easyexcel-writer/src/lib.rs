//! XLSX writer backed by `rust_xlsxwriter`.

use std::collections::HashSet;
use std::marker::PhantomData;
use std::path::{Path, PathBuf};

use easyexcel_core::{
    CellValue, ExcelColumn, ExcelError, ExcelRow, Result, WriteCellContext, WriteHandler,
    WriteRowContext, WriteSheetContext, WriteWorkbookContext,
};
use rust_xlsxwriter::{Format, FormatAlign, FormatPattern, Image, Note, Workbook, Worksheet};

/// Horizontal cell alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalAlignment {
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
}

/// Vertical cell alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
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

/// Backend-neutral write style for headers or content rows.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellStyle {
    /// Bold font.
    pub bold: bool,
    /// Italic font.
    pub italic: bool,
    /// RGB font color, for example `0xFF0000`.
    pub font_color: Option<u32>,
    /// Solid RGB background color.
    pub background_color: Option<u32>,
    /// Horizontal alignment.
    pub horizontal_alignment: Option<HorizontalAlignment>,
    /// Vertical alignment.
    pub vertical_alignment: Option<VerticalAlignment>,
    /// Wrap cell text.
    pub wrap_text: bool,
    /// Excel number format string.
    pub number_format: Option<String>,
}

impl CellStyle {
    /// Creates an empty style.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            bold: false,
            italic: false,
            font_color: None,
            background_color: None,
            horizontal_alignment: None,
            vertical_alignment: None,
            wrap_text: false,
            number_format: None,
        }
    }

    /// Sets bold font rendering.
    #[must_use]
    pub const fn bold(mut self, enabled: bool) -> Self {
        self.bold = enabled;
        self
    }

    /// Sets italic font rendering.
    #[must_use]
    pub const fn italic(mut self, enabled: bool) -> Self {
        self.italic = enabled;
        self
    }

    /// Sets the RGB font color.
    #[must_use]
    pub const fn font_color(mut self, color: u32) -> Self {
        self.font_color = Some(color);
        self
    }

    /// Sets a solid RGB background color.
    #[must_use]
    pub const fn background_color(mut self, color: u32) -> Self {
        self.background_color = Some(color);
        self
    }

    /// Sets horizontal alignment.
    #[must_use]
    pub const fn horizontal_alignment(mut self, alignment: HorizontalAlignment) -> Self {
        self.horizontal_alignment = Some(alignment);
        self
    }

    /// Sets vertical alignment.
    #[must_use]
    pub const fn vertical_alignment(mut self, alignment: VerticalAlignment) -> Self {
        self.vertical_alignment = Some(alignment);
        self
    }

    /// Enables or disables text wrapping.
    #[must_use]
    pub const fn wrap_text(mut self, enabled: bool) -> Self {
        self.wrap_text = enabled;
        self
    }

    /// Sets an Excel number format string.
    #[must_use]
    pub fn number_format(mut self, format: impl Into<String>) -> Self {
        self.number_format = Some(format.into());
        self
    }
}

/// One absolute merged-cell range using zero-based inclusive coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MergeRange {
    /// First row.
    pub first_row: u32,
    /// Last row.
    pub last_row: u32,
    /// First column.
    pub first_column: u16,
    /// Last column.
    pub last_column: u16,
}

impl MergeRange {
    /// Creates an absolute merge range.
    #[must_use]
    pub const fn new(first_row: u32, last_row: u32, first_column: u16, last_column: u16) -> Self {
        Self {
            first_row,
            last_row,
            first_column,
            last_column,
        }
    }
}

/// Repeating merge strategy equivalent to Java `LoopMergeStrategy`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LoopMergeStrategy {
    each_rows: u32,
    column_extend: u16,
    column_index: u16,
}

impl LoopMergeStrategy {
    /// Creates a repeating merge strategy.
    ///
    /// # Errors
    ///
    /// Returns a format error when dimensions are zero or no range would be merged.
    pub fn new(each_rows: u32, column_extend: u16, column_index: u16) -> Result<Self> {
        if each_rows == 0 || column_extend == 0 {
            return Err(ExcelError::Format(
                "loop merge row and column spans must be greater than zero".to_owned(),
            ));
        }
        if each_rows == 1 && column_extend == 1 {
            return Err(ExcelError::Format(
                "loop merge must span multiple rows or columns".to_owned(),
            ));
        }
        Ok(Self {
            each_rows,
            column_extend,
            column_index,
        })
    }

    /// Number of rows in each merge group.
    #[must_use]
    pub const fn each_rows(self) -> u32 {
        self.each_rows
    }

    /// Number of columns in each merge group.
    #[must_use]
    pub const fn column_extend(self) -> u16 {
        self.column_extend
    }

    /// Zero-based first column.
    #[must_use]
    pub const fn column_index(self) -> u16 {
        self.column_index
    }
}

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
    /// Absolute ranges merged before row data is written.
    pub merge_ranges: Vec<MergeRange>,
    /// Whether used columns are auto-fitted after writing.
    pub auto_width: bool,
    /// Explicit column widths in Excel character units.
    pub column_widths: Vec<(u16, u16)>,
    /// Style applied to header cells.
    pub head_style: CellStyle,
    /// Content styles cycled by relative data-row index.
    pub content_styles: Vec<CellStyle>,
    /// Repeating merge strategies applied to data rows.
    pub loop_merges: Vec<LoopMergeStrategy>,
    /// Optional dynamic multi-level head paths, one path per selected column.
    pub dynamic_head: Option<Vec<Vec<String>>>,
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
            merge_ranges: Vec::new(),
            auto_width: false,
            column_widths: Vec::new(),
            head_style: CellStyle::new().bold(true),
            content_styles: Vec::new(),
            loop_merges: Vec::new(),
            dynamic_head: None,
        }
    }
}

/// Typed worksheet metadata used by [`ExcelWriter`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheet<T> {
    options: WriteOptions,
    marker: PhantomData<T>,
}

impl<T> WriteSheet<T> {
    /// Creates worksheet metadata with the supplied name.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            options: WriteOptions {
                sheet_name: name.into(),
                ..WriteOptions::default()
            },
            marker: PhantomData,
        }
    }

    /// Returns the effective write options.
    #[must_use]
    pub const fn options(&self) -> &WriteOptions {
        &self.options
    }

    /// Enables or disables headers for this sheet.
    #[must_use]
    pub const fn need_head(mut self, enabled: bool) -> Self {
        self.options.need_head = enabled;
        self
    }

    /// Enables or disables constant-memory output for this sheet.
    #[must_use]
    pub const fn constant_memory(mut self, enabled: bool) -> Self {
        self.options.constant_memory = enabled;
        self
    }

    /// Freezes the header row for this sheet.
    #[must_use]
    pub const fn freeze_head(mut self, enabled: bool) -> Self {
        self.options.freeze_head = enabled;
        self
    }

    /// Adds an absolute merged-cell range.
    #[must_use]
    pub fn merge_cells(mut self, range: MergeRange) -> Self {
        self.options.merge_ranges.push(range);
        self
    }

    /// Enables or disables automatic width calculation.
    #[must_use]
    pub const fn auto_width(mut self, enabled: bool) -> Self {
        self.options.auto_width = enabled;
        self
    }

    /// Sets an explicit width for a zero-based physical column.
    #[must_use]
    pub fn column_width(mut self, column: u16, width: u16) -> Self {
        self.options.column_widths.push((column, width));
        self
    }

    /// Replaces the header style.
    #[must_use]
    pub fn head_style(mut self, style: CellStyle) -> Self {
        self.options.head_style = style;
        self
    }

    /// Uses one style for every content row.
    #[must_use]
    pub fn content_style(mut self, style: CellStyle) -> Self {
        self.options.content_styles = vec![style];
        self
    }

    /// Cycles the supplied styles across content rows.
    #[must_use]
    pub fn content_styles(mut self, styles: impl IntoIterator<Item = CellStyle>) -> Self {
        self.options.content_styles = styles.into_iter().collect();
        self
    }

    /// Registers a repeating data-row merge strategy.
    #[must_use]
    pub fn loop_merge(mut self, strategy: LoopMergeStrategy) -> Self {
        self.options.loop_merges.push(strategy);
        self
    }

    /// Replaces the derived headers with dynamic multi-level head paths.
    #[must_use]
    pub fn head<S, P>(mut self, paths: impl IntoIterator<Item = P>) -> Self
    where
        S: Into<String>,
        P: IntoIterator<Item = S>,
    {
        self.options.dynamic_head = Some(
            paths
                .into_iter()
                .map(|path| path.into_iter().map(Into::into).collect())
                .collect(),
        );
        self
    }
}

/// Stateful multi-sheet XLSX writer matching Java `ExcelWriter`'s lifecycle.
pub struct ExcelWriter {
    path: PathBuf,
    workbook: Workbook,
    handlers: Vec<Box<dyn WriteHandler>>,
    sheet_names: HashSet<String>,
    started: bool,
    finished: bool,
}

impl ExcelWriter {
    /// Creates a multi-sheet writer without handlers.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        Self::with_handlers(path, Vec::new())
    }

    /// Creates a multi-sheet writer with owned lifecycle handlers.
    #[must_use]
    pub fn with_handlers(path: impl Into<PathBuf>, handlers: Vec<Box<dyn WriteHandler>>) -> Self {
        Self {
            path: path.into(),
            workbook: Workbook::new(),
            handlers,
            sheet_names: HashSet::new(),
            started: false,
            finished: false,
        }
    }

    /// Writes a batch to a worksheet. Multiple calls may target different sheets.
    ///
    /// # Errors
    ///
    /// Returns an error when the writer is finished, a handler fails, or data cannot be written.
    pub fn write<T, I>(&mut self, rows: I, sheet: &WriteSheet<T>) -> Result<&mut Self>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        if self.finished {
            return Err(ExcelError::Unsupported(
                "writer already finished".to_owned(),
            ));
        }
        self.start()?;
        let sheet_name = sheet.options().sheet_name.clone();
        if !self.sheet_names.insert(sheet_name.clone()) {
            return Err(ExcelError::Format(format!(
                "duplicate worksheet name: {sheet_name}"
            )));
        }
        let result = write_sheet_to_workbook::<T, I>(
            &mut self.workbook,
            sheet.options(),
            rows,
            &mut self.handlers,
        );
        match result {
            Ok(()) => Ok(self),
            Err(error) => {
                self.sheet_names.remove(&sheet_name);
                Err(error)
            }
        }
    }

    /// Saves and closes the writer. Repeated calls are no-ops.
    ///
    /// # Errors
    ///
    /// Returns an output or handler error.
    pub fn finish(&mut self) -> Result<()> {
        if self.finished {
            return Ok(());
        }
        self.start()?;
        self.workbook.save(&self.path).map_err(format_error)?;
        let context = WriteWorkbookContext::new(&self.path);
        after_workbook(&mut self.handlers, &context)?;
        self.finished = true;
        Ok(())
    }

    /// Returns whether [`Self::finish`] completed successfully.
    #[must_use]
    pub const fn is_finished(&self) -> bool {
        self.finished
    }

    fn start(&mut self) -> Result<()> {
        if self.started {
            return Ok(());
        }
        sort_handlers(&mut self.handlers);
        let context = WriteWorkbookContext::new(&self.path);
        before_workbook(&mut self.handlers, &context)?;
        self.started = true;
        Ok(())
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
    write_xlsx_with_handlers(path, options, rows, &mut [])
}

/// Writes typed rows while invoking ordered write handlers.
///
/// # Errors
///
/// Returns a conversion, handler, worksheet-configuration, XLSX-format, or I/O error.
pub fn write_xlsx_with_handlers<T, I>(
    path: &Path,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    sort_handlers(handlers);
    let workbook_context = WriteWorkbookContext::new(path);
    before_workbook(handlers, &workbook_context)?;

    let mut workbook = Workbook::new();
    write_sheet_to_workbook::<T, I>(&mut workbook, options, rows, handlers)?;
    workbook.save(path).map_err(format_error)?;
    after_workbook(handlers, &workbook_context)?;
    Ok(())
}

fn write_sheet_to_workbook<T, I>(
    workbook: &mut Workbook,
    options: &WriteOptions,
    rows: I,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()>
where
    T: ExcelRow,
    I: IntoIterator<Item = T>,
{
    let worksheet = if options.constant_memory {
        workbook.add_worksheet_with_constant_memory()
    } else {
        workbook.add_worksheet()
    };
    worksheet
        .set_name(&options.sheet_name)
        .map_err(format_error)?;
    for (column, width) in &options.column_widths {
        worksheet
            .set_column_width(*column, *width)
            .map_err(format_error)?;
    }
    for range in &options.merge_ranges {
        worksheet
            .merge_range(
                range.first_row,
                range.first_column,
                range.last_row,
                range.last_column,
                "",
                &Format::new(),
            )
            .map_err(format_error)?;
    }
    let head_rows = dynamic_head_rows(options)?;
    let freeze_panes = options
        .freeze_panes
        .or_else(|| (options.freeze_head && options.need_head).then_some((head_rows, 0)));
    if let Some((row, column)) = freeze_panes {
        worksheet
            .set_freeze_panes(row, column)
            .map_err(format_error)?;
    }

    let sheet_context = WriteSheetContext::new(&options.sheet_name);
    before_sheet(handlers, &sheet_context)?;

    let columns = selected_columns(T::schema(), options);
    let mut row_index = 0_u32;
    if options.need_head {
        if let Some(head) = &options.dynamic_head {
            write_dynamic_headers_with_handlers(
                worksheet,
                &columns,
                head,
                &options.sheet_name,
                &options.head_style,
                handlers,
            )?;
        } else {
            write_headers_with_handlers(
                worksheet,
                &columns,
                &options.sheet_name,
                &options.head_style,
                handlers,
            )?;
        }
        row_index = head_rows;
    }
    for (data_index, row) in rows.into_iter().enumerate() {
        let cells = row.to_row()?;
        let style = (!options.content_styles.is_empty())
            .then(|| &options.content_styles[data_index % options.content_styles.len()]);
        apply_loop_merges(worksheet, row_index, data_index, &options.loop_merges)?;
        write_data_row_with_handlers(
            worksheet,
            row_index,
            &columns,
            &cells,
            &options.sheet_name,
            style,
            handlers,
        )?;
        row_index += 1;
    }
    after_sheet(handlers, &sheet_context)?;
    if options.auto_width {
        worksheet.autofit();
    }
    Ok(())
}

fn apply_loop_merges(
    worksheet: &mut Worksheet,
    row_index: u32,
    data_index: usize,
    strategies: &[LoopMergeStrategy],
) -> Result<()> {
    for strategy in strategies {
        #[allow(clippy::cast_possible_truncation)]
        let each_rows = strategy.each_rows as usize;
        if !data_index.is_multiple_of(each_rows) {
            continue;
        }
        let last_row = row_index
            .checked_add(strategy.each_rows - 1)
            .ok_or_else(|| ExcelError::Format("loop merge row overflow".to_owned()))?;
        let last_column = strategy
            .column_index
            .checked_add(strategy.column_extend - 1)
            .ok_or_else(|| ExcelError::Format("loop merge column overflow".to_owned()))?;
        worksheet
            .merge_range(
                row_index,
                strategy.column_index,
                last_row,
                last_column,
                "",
                &Format::new(),
            )
            .map_err(format_error)?;
    }
    Ok(())
}

fn sort_handlers(handlers: &mut [Box<dyn WriteHandler>]) {
    handlers.sort_by_key(|handler| handler.order());
}

fn before_workbook(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.before_workbook(context)?;
    }
    Ok(())
}

fn after_workbook(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.after_workbook(context)?;
    }
    Ok(())
}

fn before_sheet(handlers: &mut [Box<dyn WriteHandler>], context: &WriteSheetContext) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.before_sheet(context)?;
    }
    Ok(())
}

fn after_sheet(handlers: &mut [Box<dyn WriteHandler>], context: &WriteSheetContext) -> Result<()> {
    for handler in handlers.iter_mut() {
        handler.after_sheet(context)?;
    }
    Ok(())
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

fn dynamic_head_rows(options: &WriteOptions) -> Result<u32> {
    if !options.need_head {
        return Ok(0);
    }
    let Some(head) = &options.dynamic_head else {
        return Ok(1);
    };
    if head.is_empty() || head.iter().any(Vec::is_empty) {
        return Err(ExcelError::Format(
            "dynamic head must contain at least one non-empty path".to_owned(),
        ));
    }
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    head_level_to_row(levels)
}

fn head_level_to_row(level: usize) -> Result<u32> {
    u32::try_from(level).map_err(|_| ExcelError::Format("dynamic head is too deep".to_owned()))
}

#[cfg(test)]
fn write_headers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
) -> Result<()> {
    write_headers_with_handlers(worksheet, columns, "", &CellStyle::default(), &mut [])
}

fn write_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    sheet_name: &str,
    style: &CellStyle,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    let labels = columns
        .iter()
        .map(|(_, _, column)| column.name.to_owned())
        .collect::<Vec<_>>();
    write_header_row_with_handlers(worksheet, 0, columns, &labels, sheet_name, style, handlers)
}

fn write_dynamic_headers_with_handlers(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    head: &[Vec<String>],
    sheet_name: &str,
    style: &CellStyle,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    if head.len() != columns.len() {
        return Err(ExcelError::Format(format!(
            "dynamic head column count {} does not match selected column count {}",
            head.len(),
            columns.len()
        )));
    }
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    for level in 0..levels {
        #[allow(clippy::cast_possible_truncation)]
        let row_index = level as u32;
        let labels = head
            .iter()
            .map(|path| path.get(level).cloned().unwrap_or_default())
            .collect::<Vec<_>>();
        write_header_row_with_handlers(
            worksheet, row_index, columns, &labels, sheet_name, style, handlers,
        )?;
    }
    merge_dynamic_head_groups(worksheet, columns, head, style)
}

fn write_header_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    labels: &[String],
    sheet_name: &str,
    style: &CellStyle,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    let format = cell_format(Some(style));
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: true,
    };
    for handler in handlers.iter_mut() {
        handler.before_row(&row_context)?;
    }
    for ((physical_index, _, column), label) in columns.iter().zip(labels) {
        let column_index = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index,
            field: Some(column.field),
            is_head: true,
            value: CellValue::String(label.clone()),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            match &context.value {
                CellValue::String(value) | CellValue::Error(value) => {
                    worksheet
                        .write_string_with_format(row_index, context.column_index, value, &format)
                        .map_err(format_error)?;
                }
                value => write_cell(
                    worksheet,
                    row_index,
                    context.column_index,
                    column,
                    value,
                    Some(style),
                )?,
            }
        }
        for handler in handlers.iter_mut() {
            handler.after_cell(&context)?;
        }
    }
    for handler in handlers.iter_mut() {
        handler.after_row(&row_context)?;
    }
    Ok(())
}

fn merge_dynamic_head_groups(
    worksheet: &mut Worksheet,
    columns: &[(usize, usize, &'static ExcelColumn)],
    head: &[Vec<String>],
    style: &CellStyle,
) -> Result<()> {
    let levels = head.iter().map(Vec::len).max().unwrap_or(0);
    let format = cell_format(Some(style));
    for level in 0..levels {
        #[allow(clippy::cast_possible_truncation)]
        let row_index = level as u32;
        let mut start = 0;
        while start < head.len() {
            let mut end = start;
            while end + 1 < head.len()
                && columns[end].0.checked_add(1) == Some(columns[end + 1].0)
                && same_dynamic_head_group(head, start, end + 1, level)
            {
                end += 1;
            }
            let label = head[start].get(level).map_or("", String::as_str);
            if end > start && !label.is_empty() {
                worksheet
                    .merge_range(
                        row_index,
                        to_column(columns[start].0)?,
                        row_index,
                        to_column(columns[end].0)?,
                        label,
                        &format,
                    )
                    .map_err(format_error)?;
            }
            start = end + 1;
        }
    }
    Ok(())
}

fn same_dynamic_head_group(
    head: &[Vec<String>],
    first: usize,
    second: usize,
    level: usize,
) -> bool {
    head[first].get(level) == head[second].get(level)
        && head[first]
            .iter()
            .take(level)
            .eq(head[second].iter().take(level))
}

#[cfg(test)]
fn write_data_row(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
) -> Result<()> {
    write_data_row_with_handlers(worksheet, row_index, columns, cells, "", None, &mut [])
}

fn write_data_row_with_handlers(
    worksheet: &mut Worksheet,
    row_index: u32,
    columns: &[(usize, usize, &'static ExcelColumn)],
    cells: &[CellValue],
    sheet_name: &str,
    style: Option<&CellStyle>,
    handlers: &mut [Box<dyn WriteHandler>],
) -> Result<()> {
    let row_context = WriteRowContext {
        sheet_name: sheet_name.to_owned(),
        row_index,
        is_head: false,
    };
    for handler in handlers.iter_mut() {
        handler.before_row(&row_context)?;
    }
    for (physical_index, schema_index, metadata) in columns {
        let value = cells.get(*schema_index).unwrap_or(&CellValue::Empty);
        let column = to_column(*physical_index)?;
        let mut context = WriteCellContext {
            sheet_name: sheet_name.to_owned(),
            row_index,
            column_index: column,
            field: Some(metadata.field),
            is_head: false,
            value: value.clone(),
            skip: false,
        };
        for handler in handlers.iter_mut() {
            handler.before_cell(&mut context)?;
        }
        if !context.skip {
            write_cell(
                worksheet,
                row_index,
                context.column_index,
                metadata,
                &context.value,
                style,
            )?;
        }
        for handler in handlers.iter_mut() {
            handler.after_cell(&context)?;
        }
    }
    for handler in handlers.iter_mut() {
        handler.after_row(&row_context)?;
    }
    Ok(())
}

fn write_cell(
    worksheet: &mut Worksheet,
    row_index: u32,
    column: u16,
    metadata: &ExcelColumn,
    value: &CellValue,
    style: Option<&CellStyle>,
) -> Result<()> {
    let format = cell_format(style);
    match value {
        CellValue::Empty => {
            worksheet
                .write_blank(row_index, column, &format)
                .map_err(format_error)?;
        }
        CellValue::String(value) | CellValue::Error(value) => {
            worksheet
                .write_string_with_format(row_index, column, value, &format)
                .map_err(format_error)?;
        }
        CellValue::Bool(value) => {
            worksheet
                .write_boolean_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::Int(value) => {
            write_integer(worksheet, row_index, column, *value, &format)?;
        }
        CellValue::Float(value) => {
            worksheet
                .write_number_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::Date(value) => {
            let format = format
                .clone()
                .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd"));
            worksheet
                .write_datetime_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::DateTime(value) => {
            let format = format
                .clone()
                .set_num_format(excel_date_format(metadata.format, "yyyy-mm-dd hh:mm:ss"));
            worksheet
                .write_datetime_with_format(row_index, column, *value, &format)
                .map_err(format_error)?;
        }
        CellValue::Formula(value) => {
            worksheet
                .write_formula_with_format(row_index, column, value.as_str(), &format)
                .map_err(format_error)?;
        }
        CellValue::Hyperlink { url, text } => {
            worksheet
                .write_url_with_options(row_index, column, url.as_str(), text, "", Some(&format))
                .map_err(format_error)?;
        }
        CellValue::Comment { value, text } => {
            write_cell(worksheet, row_index, column, metadata, value, style)?;
            worksheet
                .insert_note(row_index, column, &Note::new(text))
                .map_err(format_error)?;
        }
        CellValue::Image(bytes) => {
            let image = image_from_buffer(bytes)?;
            worksheet
                .insert_image_fit_to_cell(row_index, column, &image, true)
                .map_err(format_error)?;
        }
    }
    Ok(())
}

fn image_from_buffer(bytes: &[u8]) -> Result<Image> {
    if bytes.len() < 8 {
        return Err(ExcelError::Format(
            "image buffer is too short to contain a valid header".to_owned(),
        ));
    }
    Image::new_from_buffer(bytes).map_err(format_error)
}

fn cell_format(style: Option<&CellStyle>) -> Format {
    let Some(style) = style else {
        return Format::new();
    };
    let mut format = Format::new();
    if style.bold {
        format = format.set_bold();
    }
    if style.italic {
        format = format.set_italic();
    }
    if let Some(color) = style.font_color {
        format = format.set_font_color(color);
    }
    if let Some(color) = style.background_color {
        format = format
            .set_background_color(color)
            .set_pattern(FormatPattern::Solid);
    }
    if let Some(alignment) = style.horizontal_alignment {
        format = format.set_align(horizontal_format_align(alignment));
    }
    if let Some(alignment) = style.vertical_alignment {
        format = format.set_align(vertical_format_align(alignment));
    }
    if style.wrap_text {
        format = format.set_text_wrap();
    }
    if let Some(number_format) = &style.number_format {
        format = format.set_num_format(number_format);
    }
    format
}

const fn horizontal_format_align(alignment: HorizontalAlignment) -> FormatAlign {
    match alignment {
        HorizontalAlignment::General => FormatAlign::General,
        HorizontalAlignment::Left => FormatAlign::Left,
        HorizontalAlignment::Center => FormatAlign::Center,
        HorizontalAlignment::Right => FormatAlign::Right,
        HorizontalAlignment::Fill => FormatAlign::Fill,
        HorizontalAlignment::Justify => FormatAlign::Justify,
        HorizontalAlignment::CenterAcross => FormatAlign::CenterAcross,
    }
}

const fn vertical_format_align(alignment: VerticalAlignment) -> FormatAlign {
    match alignment {
        VerticalAlignment::Top => FormatAlign::Top,
        VerticalAlignment::Center => FormatAlign::VerticalCenter,
        VerticalAlignment::Bottom => FormatAlign::Bottom,
        VerticalAlignment::Justify => FormatAlign::VerticalJustify,
        VerticalAlignment::Distributed => FormatAlign::VerticalDistributed,
    }
}

fn write_integer(
    worksheet: &mut Worksheet,
    row: u32,
    column: u16,
    value: i64,
    format: &Format,
) -> Result<()> {
    const MAX_EXACT_EXCEL_INTEGER: u64 = 9_007_199_254_740_991;
    if value.unsigned_abs() <= MAX_EXACT_EXCEL_INTEGER {
        #[allow(clippy::cast_precision_loss)]
        let number = value as f64;
        worksheet
            .write_number_with_format(row, column, number, format)
            .map_err(format_error)?;
    } else {
        worksheet
            .write_string_with_format(row, column, value.to_string(), format)
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
