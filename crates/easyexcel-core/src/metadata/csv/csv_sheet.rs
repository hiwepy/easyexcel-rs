//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvSheet`.

use std::collections::VecDeque;

use crate::excel_error::ExcelError;
use crate::util::work_book_util::RowCreator;

use super::csv_row::CsvRow;

/// Single-sheet, ordered-row CSV model.
#[derive(Debug, Clone, PartialEq)]
pub struct CsvSheet {
    name: String,
    row_cache_count: usize,
    last_row_index: Option<u32>,
    row_cache: VecDeque<CsvRow>,
}

impl CsvSheet {
    /// Creates an empty sheet with Java's default 100-row cache.
    #[must_use]
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            row_cache_count: 100,
            last_row_index: None,
            row_cache: VecDeque::with_capacity(100),
        }
    }

    /// Returns the logical sheet name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Sets the expected first row for a stateful append.
    pub fn set_next_row_index(&mut self, next_row_index: u32) {
        self.last_row_index = next_row_index.checked_sub(1);
    }

    /// Returns the last created row index.
    #[must_use]
    pub const fn last_row_index(&self) -> Option<u32> {
        self.last_row_index
    }

    /// Returns a cached row, or an error after it has been flushed.
    pub fn row(&self, row_index: u32) -> Result<&CsvRow, ExcelError> {
        self.row_cache
            .iter()
            .find(|row| row.row_index() == row_index)
            .ok_or_else(|| {
                ExcelError::Unsupported("the CSV row does not exist or has been flushed".to_owned())
            })
    }

    /// Removes and returns the most recently created row.
    pub fn take_last_row(&mut self) -> Option<CsvRow> {
        self.row_cache.pop_back()
    }

    /// Returns rows that exceed the configured cache size.
    pub fn drain_flushable_rows(&mut self) -> Vec<CsvRow> {
        let count = self.row_cache.len().saturating_sub(self.row_cache_count);
        self.row_cache.drain(..count).collect()
    }
}

impl RowCreator for CsvSheet {
    type Row<'a>
        = &'a mut CsvRow
    where
        Self: 'a;

    fn create_row(&mut self, row_index: u32) -> Result<Self::Row<'_>, ExcelError> {
        let expected = self
            .last_row_index
            .map_or(0, |last_row_index| last_row_index.saturating_add(1));
        if row_index != expected {
            return Err(ExcelError::Format(format!(
                "CSV rows must be created in order: expected {expected}, got {row_index}"
            )));
        }
        self.last_row_index = Some(row_index);
        self.row_cache.push_back(CsvRow::new(row_index));
        Ok(self.row_cache.back_mut().expect("just pushed"))
    }
}
