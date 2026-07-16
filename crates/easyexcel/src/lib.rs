//! Public facade for typed, event-driven Excel reading and writing.

use std::marker::PhantomData;
use std::path::{Path, PathBuf};

pub use easyexcel_core::*;
pub use easyexcel_derive::ExcelRow;
use easyexcel_reader::{ReadOptions, SheetSelector, read_xlsx};
use easyexcel_writer::{WriteOptions, write_xlsx};

/// Static factory matching Java `EasyExcel`'s entry point.
pub struct EasyExcel;

impl EasyExcel {
    /// Starts an event-driven XLSX read.
    pub fn read<T, L>(path: impl Into<PathBuf>, listener: L) -> ExcelReaderBuilder<T, L>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        ExcelReaderBuilder {
            path: path.into(),
            options: ReadOptions::default(),
            listener,
            marker: PhantomData,
        }
    }

    /// Starts a synchronous read that collects all converted rows.
    pub fn read_sync<T>(path: impl Into<PathBuf>) -> ExcelSyncReaderBuilder<T>
    where
        T: ExcelRow,
    {
        ExcelSyncReaderBuilder {
            path: path.into(),
            options: ReadOptions::default(),
            marker: PhantomData,
        }
    }

    /// Starts a new XLSX write.
    pub fn write<T>(path: impl Into<PathBuf>) -> ExcelWriterBuilder<T>
    where
        T: ExcelRow,
    {
        ExcelWriterBuilder {
            path: path.into(),
            options: WriteOptions::default(),
            marker: PhantomData,
        }
    }
}

/// Input accepted by `.sheet(...)`.
pub trait IntoSheetSelector {
    /// Converts to internal sheet selection.
    fn into_sheet_selector(self) -> SheetSelector;
}

impl IntoSheetSelector for usize {
    fn into_sheet_selector(self) -> SheetSelector {
        SheetSelector::Index(self)
    }
}

impl IntoSheetSelector for &str {
    fn into_sheet_selector(self) -> SheetSelector {
        SheetSelector::Name(self.to_owned())
    }
}

impl IntoSheetSelector for String {
    fn into_sheet_selector(self) -> SheetSelector {
        SheetSelector::Name(self)
    }
}

/// Event-driven reader builder.
pub struct ExcelReaderBuilder<T, L> {
    path: PathBuf,
    options: ReadOptions,
    listener: L,
    marker: PhantomData<T>,
}

impl<T, L> ExcelReaderBuilder<T, L>
where
    T: ExcelRow,
    L: ReadListener<T>,
{
    /// Selects a worksheet by name or zero-based index.
    #[must_use]
    pub fn sheet(mut self, sheet: impl IntoSheetSelector) -> Self {
        self.options.sheet = sheet.into_sheet_selector();
        self
    }

    /// Selects every worksheet in workbook order.
    #[must_use]
    pub fn all_sheets(mut self) -> Self {
        self.options.sheet = SheetSelector::All;
        self
    }

    /// Sets the number of header rows.
    #[must_use]
    pub const fn head_row_number(mut self, rows: u32) -> Self {
        self.options.head_row_number = rows;
        self
    }

    /// Configures empty-row filtering.
    #[must_use]
    pub const fn ignore_empty_row(mut self, ignore: bool) -> Self {
        self.options.ignore_empty_row = ignore;
        self
    }

    /// Executes the read and consumes the builder.
    ///
    /// # Errors
    ///
    /// Returns a workbook, sheet-selection, conversion, or listener error.
    pub fn do_read(mut self) -> Result<()> {
        read_xlsx::<T, L>(&self.path, &self.options, &mut self.listener)
    }
}

/// Synchronous collecting reader builder.
pub struct ExcelSyncReaderBuilder<T> {
    path: PathBuf,
    options: ReadOptions,
    marker: PhantomData<T>,
}

impl<T> ExcelSyncReaderBuilder<T>
where
    T: ExcelRow,
{
    /// Selects a worksheet by name or zero-based index.
    #[must_use]
    pub fn sheet(mut self, sheet: impl IntoSheetSelector) -> Self {
        self.options.sheet = sheet.into_sheet_selector();
        self
    }

    /// Sets the number of header rows.
    #[must_use]
    pub const fn head_row_number(mut self, rows: u32) -> Self {
        self.options.head_row_number = rows;
        self
    }

    /// Reads all rows into memory.
    ///
    /// # Errors
    ///
    /// Returns a workbook, sheet-selection, or row-conversion error.
    pub fn do_read_sync(self) -> Result<Vec<T>> {
        let mut listener = CollectListener(Vec::new());
        read_xlsx::<T, _>(&self.path, &self.options, &mut listener)?;
        Ok(listener.0)
    }
}

struct CollectListener<T>(Vec<T>);

impl<T> ReadListener<T> for CollectListener<T> {
    fn invoke(&mut self, data: T, _context: &AnalysisContext) -> Result<()> {
        self.0.push(data);
        Ok(())
    }
}

/// New-workbook writer builder.
pub struct ExcelWriterBuilder<T> {
    path: PathBuf,
    options: WriteOptions,
    marker: PhantomData<T>,
}

impl<T> ExcelWriterBuilder<T>
where
    T: ExcelRow,
{
    /// Sets the worksheet name.
    #[must_use]
    pub fn sheet(mut self, name: impl Into<String>) -> Self {
        self.options.sheet_name = name.into();
        self
    }

    /// Enables or disables the header row.
    #[must_use]
    pub const fn need_head(mut self, need_head: bool) -> Self {
        self.options.need_head = need_head;
        self
    }

    /// Freezes the header row.
    #[must_use]
    pub const fn freeze_head(mut self, freeze: bool) -> Self {
        self.options.freeze_head = freeze;
        self
    }

    /// Freezes rows and columns above and to the left of the position.
    #[must_use]
    pub const fn freeze_panes(mut self, row: u32, column: u16) -> Self {
        self.options.freeze_panes = Some((row, column));
        self
    }

    /// Includes only the supplied physical column indexes.
    #[must_use]
    pub fn include_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        self.options.include_column_indexes = Some(indexes.into_iter().collect());
        self
    }

    /// Includes only the supplied Rust field names.
    #[must_use]
    pub fn include_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        self.options.include_column_field_names = Some(names.into_iter().map(Into::into).collect());
        self
    }

    /// Excludes physical column indexes.
    #[must_use]
    pub fn exclude_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        self.options.exclude_column_indexes = indexes.into_iter().collect();
        self
    }

    /// Excludes Rust field names.
    #[must_use]
    pub fn exclude_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        self.options.exclude_column_field_names = names.into_iter().map(Into::into).collect();
        self
    }

    /// Orders selected columns by the corresponding include list.
    #[must_use]
    pub const fn order_by_include_column(mut self, enabled: bool) -> Self {
        self.options.order_by_include_column = enabled;
        self
    }

    /// Selects constant-memory output.
    #[must_use]
    pub const fn constant_memory(mut self, enabled: bool) -> Self {
        self.options.constant_memory = enabled;
        self
    }

    /// Writes any owned row iterator.
    ///
    /// # Errors
    ///
    /// Returns a conversion, worksheet-configuration, XLSX-format, or I/O error.
    pub fn do_write<I>(self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        write_xlsx::<T, I>(Path::new(&self.path), &self.options, rows)
    }

    /// Alias emphasizing that the input is consumed incrementally.
    ///
    /// # Errors
    ///
    /// Returns a conversion, worksheet-configuration, XLSX-format, or I/O error.
    pub fn do_write_iter<I>(self, rows: I) -> Result<()>
    where
        I: IntoIterator<Item = T>,
    {
        self.do_write(rows)
    }
}

#[cfg(test)]
mod tests;
