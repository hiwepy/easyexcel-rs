//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvCellStyle`.

use crate::metadata::data::DataFormatData;

/// CSV cell-style metadata.
///
/// Like Java, only the style index and data format affect CSV rendering.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CsvCellStyle {
    index: i16,
    data_format_data: Option<DataFormatData>,
}

impl CsvCellStyle {
    /// Creates a style with its workbook-local index.
    #[must_use]
    pub const fn new(index: i16) -> Self {
        Self {
            index,
            data_format_data: None,
        }
    }

    /// Returns the workbook-local style index.
    #[must_use]
    pub const fn index(&self) -> i16 {
        self.index
    }

    /// Sets the numeric data-format index.
    pub fn set_data_format(&mut self, format: i16) {
        self.data_format_data
            .get_or_insert_with(DataFormatData::default)
            .index = Some(format);
    }

    /// Sets an owned data-format string.
    pub fn set_data_format_string(&mut self, format: impl Into<String>) {
        self.data_format_data
            .get_or_insert_with(DataFormatData::default)
            .format = Some(format.into());
    }

    /// Returns the nested data-format metadata.
    #[must_use]
    pub const fn data_format_data(&self) -> Option<&DataFormatData> {
        self.data_format_data.as_ref()
    }
}
