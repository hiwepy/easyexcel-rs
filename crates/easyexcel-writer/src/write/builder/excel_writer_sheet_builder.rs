//! Rust implementation of Java
//! `com.alibaba.excel.write.builder.ExcelWriterSheetBuilder`.

use easyexcel_core::{ExcelError, ExcelRow, Result, WriteHandler};

use crate::metadata::write_sheet::WriteSheet as WriteSheetMetadata;
use crate::{ExcelWriter, ExcelWriterTableBuilder, WriteOptions, WriteSheet};

/// A sheet builder optionally owning the writer that will execute it.
pub struct ExcelWriterSheetBuilder {
    excel_writer: Option<ExcelWriter>,
    write_sheet: WriteSheetMetadata,
    handlers: Vec<Box<dyn WriteHandler>>,
}

impl ExcelWriterSheetBuilder {
    /// Creates an unbound metadata builder.
    #[must_use]
    pub fn new() -> Self {
        Self {
            excel_writer: None,
            write_sheet: WriteSheetMetadata::new(),
            handlers: Vec::new(),
        }
    }

    /// Creates a sheet builder owning its stateful writer.
    #[must_use]
    pub fn with_excel_writer(excel_writer: ExcelWriter) -> Self {
        Self {
            excel_writer: Some(excel_writer),
            write_sheet: WriteSheetMetadata::new(),
            handlers: Vec::new(),
        }
    }

    /// Creates a sheet builder with effective workbook options already
    /// inherited, while retaining nullable sheet-level overrides separately.
    #[must_use]
    pub fn with_excel_writer_and_options(
        excel_writer: ExcelWriter,
        inherited_options: WriteOptions,
    ) -> Self {
        let mut write_sheet = WriteSheetMetadata::new();
        write_sheet.options = inherited_options;
        write_sheet.sheet_no = write_sheet.options.sheet_index.unwrap_or(0) as i32;
        write_sheet.sheet_name = write_sheet.options.sheet_name.clone();
        Self {
            excel_writer: Some(excel_writer),
            write_sheet,
            handlers: Vec::new(),
        }
    }

    /// Sets the zero-based sheet number.
    #[must_use]
    pub fn sheet_no(mut self, sheet_no: i32) -> Self {
        self.write_sheet.set_sheet_no(sheet_no);
        self.write_sheet.options.sheet_index = Some(sheet_no.max(0) as usize);
        self
    }

    /// Sets the sheet name.
    #[must_use]
    pub fn sheet_name(mut self, sheet_name: impl Into<String>) -> Self {
        let sheet_name = sheet_name.into();
        self.write_sheet.set_sheet_name(sheet_name.clone());
        self.write_sheet.options.sheet_name = sheet_name;
        self
    }

    /// Sets the number of rows before the header.
    #[must_use]
    pub fn relative_head_row_index(mut self, index: i32) -> Self {
        self.write_sheet.parameter.relative_head_row_index = Some(index);
        self.write_sheet.options.relative_head_row_index = index;
        self
    }

    /// Controls header output.
    #[must_use]
    pub fn need_head(mut self, enabled: bool) -> Self {
        self.write_sheet.parameter.need_head = Some(enabled);
        self.write_sheet.options.need_head = enabled;
        self
    }

    /// Enables or disables Java's default bold header style.
    #[must_use]
    pub fn use_default_style(mut self, enabled: bool) -> Self {
        self.write_sheet.parameter.use_default_style = Some(enabled);
        self.write_sheet.options.use_default_style = enabled;
        self.write_sheet.options.head_style = if enabled {
            crate::CellStyle::new().bold(true)
        } else {
            crate::CellStyle::new()
        };
        self
    }

    /// Controls automatic multi-level header merging.
    #[must_use]
    pub fn automatic_merge_head(mut self, enabled: bool) -> Self {
        self.write_sheet.parameter.automatic_merge_head = Some(enabled);
        self.write_sheet.options.automatic_merge_head = enabled;
        self
    }

    /// Includes only the supplied physical columns.
    #[must_use]
    pub fn include_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        let indexes = indexes.into_iter().collect::<Vec<_>>();
        self.write_sheet.parameter.include_column_indexes = Some(indexes.clone());
        self.write_sheet.options.include_column_indexes = Some(indexes);
        self
    }

    /// Includes only the supplied Rust field names.
    #[must_use]
    pub fn include_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        let names = names.into_iter().map(Into::into).collect::<Vec<_>>();
        self.write_sheet.parameter.include_column_field_names = Some(names.clone());
        self.write_sheet.options.include_column_field_names = Some(names);
        self
    }

    /// Excludes physical columns.
    #[must_use]
    pub fn exclude_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        let indexes = indexes.into_iter().collect::<Vec<_>>();
        self.write_sheet.parameter.exclude_column_indexes = Some(indexes.clone());
        self.write_sheet.options.exclude_column_indexes = indexes;
        self
    }

    /// Excludes Rust field names.
    #[must_use]
    pub fn exclude_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        let names = names.into_iter().map(Into::into).collect::<Vec<_>>();
        self.write_sheet.parameter.exclude_column_field_names = Some(names.clone());
        self.write_sheet.options.exclude_column_field_names = names;
        self
    }

    /// Orders output by the include-list order.
    #[must_use]
    pub fn order_by_include_column(mut self, enabled: bool) -> Self {
        self.write_sheet.parameter.order_by_include_column = Some(enabled);
        self.write_sheet.options.order_by_include_column = enabled;
        self
    }

    /// Stores a handler owned by the sheet holder.
    #[must_use]
    pub fn register_write_handler(mut self, handler: impl WriteHandler + 'static) -> Self {
        self.handlers.push(Box::new(handler));
        self
    }

    /// Builds the untyped Java-compatible sheet metadata.
    #[must_use]
    pub fn build(&self) -> WriteSheetMetadata {
        self.write_sheet.clone()
    }

    /// Writes the supplied rows and finishes the owned writer.
    ///
    /// This mirrors Java `ExcelWriterSheetBuilder.doWrite(Collection)`.
    pub fn do_write<T, I>(mut self, rows: I) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let mut writer = self.excel_writer.take().ok_or_else(|| {
            ExcelError::Format("Must use ExcelWriterBuilder.sheet() to call do_write()".to_owned())
        })?;
        let sheet = WriteSheet::<T>::from_options(self.write_sheet.options.clone());
        writer.write_with_sheet_handlers(rows, &sheet, self.handlers)?;
        writer.finish()
    }

    /// Resolves rows lazily, then delegates to [`Self::do_write`].
    pub fn do_write_with<T, I, F>(self, supplier: F) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
        F: FnOnce() -> I,
    {
        self.do_write(supplier())
    }

    /// Creates a table builder bound to this writer and sheet.
    ///
    /// Mirrors Java `ExcelWriterSheetBuilder.table()`.
    #[must_use]
    pub fn table(mut self) -> ExcelWriterTableBuilder {
        match self.excel_writer.take() {
            Some(writer) => {
                ExcelWriterTableBuilder::with_excel_writer(writer, self.write_sheet, self.handlers)
            }
            None => ExcelWriterTableBuilder::new(),
        }
    }

    /// Creates a numbered table builder.
    ///
    /// Mirrors Java `ExcelWriterSheetBuilder.table(Integer)`.
    #[must_use]
    pub fn table_no(self, table_no: i32) -> ExcelWriterTableBuilder {
        self.table().table_no(table_no)
    }
}

impl Default for ExcelWriterSheetBuilder {
    fn default() -> Self {
        Self::new()
    }
}
