//! Mirrors Java `com.alibaba.excel.write.ExcelBuilder` and `ExcelBuilderImpl`.

use std::any::Any;
use std::path::PathBuf;

use easyexcel_core::{
    DynamicRow, ExcelError, ExcelRow, Result, WriteContext, WriteContextImpl, WriteFillConfig,
    WriteFillExecutor, WriteFillSheet, WriteSheetContext, csv_fill_unsupported_error,
    fill_requires_template_error,
};

use crate::builder::excel_writer_table_builder::merge_table_options;
use crate::metadata::WriteTable;
use crate::{ExcelWriter, MergeRange, WriteOptions, WriteSheet};

/// Minimal fill configuration accepted by [`ExcelBuilder::fill`].
///
/// Mirrors Java `com.alibaba.excel.write.metadata.fill.FillConfig` at the
/// builder surface. Stateful template filling remains on
/// `easyexcel_template::FillConfig`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct FillConfig {
    /// Whether collection fill forces a new row. (Java `FillConfig.forceNewRow`)
    pub force_new_row: bool,
}

impl FillConfig {
    /// Creates Java-compatible defaults (`forceNewRow = false`).
    #[must_use]
    pub const fn new() -> Self {
        Self {
            force_new_row: false,
        }
    }
}

/// Workbook builder contract matching Java `ExcelBuilder`.
///
/// Mirrors Java `com.alibaba.excel.write.ExcelBuilder`.
pub trait ExcelBuilder {
    /// Appends rows to a worksheet. (Java `addContent(Collection, WriteSheet)`)
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or I/O error from the underlying writer.
    fn add_content<T, I>(&mut self, data: I, write_sheet: &WriteSheet<T>) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>;

    /// Appends rows to a worksheet table. (Java `addContent(Collection, WriteSheet, WriteTable)`)
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or I/O error from the underlying writer.
    fn add_content_with_table<T, I>(
        &mut self,
        data: I,
        write_sheet: &WriteSheet<T>,
        write_table: &WriteTable,
    ) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>;

    /// Fills template placeholders on a worksheet. (Java `fill(Object, FillConfig, WriteSheet)`)
    ///
    /// `data` must be a supported fill payload (`TemplateData`, `FillWrapper`, …)
    /// wired through [`WriteFillExecutor`] by the `easyexcel` facade when a
    /// template is configured.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Unsupported`] when no template stream is configured.
    fn fill(
        &mut self,
        _data: &dyn Any,
        _fill_config: FillConfig,
        _write_sheet: &WriteSheet<DynamicRow>,
    ) -> Result<()> {
        Err(fill_requires_template_error())
    }

    /// Creates a merged region using zero-based inclusive coordinates.
    ///
    /// Mirrors deprecated Java `merge(int, int, int, int)`.
    fn merge(
        &mut self,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()>;

    /// Returns the active write context. (Java `writeContext()`)
    fn write_context(&self) -> &dyn WriteContext;

    /// Completes the workbook lifecycle. (Java `finish(boolean onException)`)
    ///
    /// # Errors
    ///
    /// Returns an output, close, or handler error.
    fn finish(&mut self, on_exception: bool) -> Result<()>;
}

/// Concrete builder implementation delegating to [`ExcelWriter`].
///
/// Mirrors Java `com.alibaba.excel.write.ExcelBuilderImpl`.
pub struct ExcelBuilderImpl {
    writer: ExcelWriter,
    logical_path: PathBuf,
    pending_merges: Vec<MergeRange>,
    context: WriteContextImpl,
    fill_executor: Option<Box<dyn WriteFillExecutor>>,
    finished_via_fill: bool,
    fill_session_active: bool,
}

impl ExcelBuilderImpl {
    /// Creates a builder from a stateful writer. (Java `new ExcelBuilderImpl(WriteWorkbook)`)
    #[must_use]
    pub fn new(writer: ExcelWriter, logical_path: impl Into<PathBuf>) -> Self {
        let logical_path = logical_path.into();
        Self {
            context: WriteContextImpl::new(&logical_path),
            writer,
            logical_path,
            pending_merges: Vec::new(),
            fill_executor: None,
            finished_via_fill: false,
            fill_session_active: false,
        }
    }

    /// Creates a builder from path and options via [`ExcelWriter::with_handlers_and_options`].
    #[must_use]
    pub fn from_options(path: impl Into<PathBuf>, options: WriteOptions) -> Self {
        let logical_path = path.into();
        Self::new(
            ExcelWriter::with_handlers_and_options(&logical_path, Vec::new(), options),
            logical_path,
        )
    }

    /// Returns the underlying writer for Java-style `ExcelWriter` facades.
    #[must_use]
    pub fn into_writer(self) -> ExcelWriter {
        self.writer
    }

    /// Returns a mutable reference to the underlying writer.
    pub fn writer_mut(&mut self) -> &mut ExcelWriter {
        &mut self.writer
    }

    /// Returns the logical output path carried by this builder.
    #[must_use]
    pub fn logical_path(&self) -> &std::path::Path {
        &self.logical_path
    }

    /// Installs a template fill executor wired by the `easyexcel` facade.
    ///
    /// Mirrors Java lazy `ExcelWriteFillExecutor` creation inside
    /// `ExcelBuilderImpl.fill`.
    pub fn set_fill_executor(&mut self, executor: Box<dyn WriteFillExecutor>) {
        self.fill_executor = Some(executor);
    }

    /// Returns whether a template fill executor has been installed.
    #[must_use]
    pub fn has_fill_executor(&self) -> bool {
        self.fill_executor.is_some()
    }

    /// Returns whether [`Self::finish`] already persisted fill output.
    #[must_use]
    pub const fn finished_via_fill(&self) -> bool {
        self.finished_via_fill
    }

    fn write_rows<T, I>(
        &mut self,
        data: I,
        write_sheet: &WriteSheet<T>,
        write_table: Option<&WriteTable>,
    ) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        let mut options = if let Some(table) = write_table {
            merge_table_options(write_sheet.options(), table)
        } else {
            write_sheet.options().clone()
        };
        options.merge_ranges.extend(self.pending_merges.drain(..));
        let sheet_name = if options.auto_trim {
            options.sheet_name.trim().to_owned()
        } else {
            options.sheet_name.clone()
        };
        options.sheet_name = sheet_name.clone();
        self.context.set_sheet_context(&sheet_name);
        self.context
            .set_table_no(write_table.map(WriteTable::table_no));
        let sheet = WriteSheet::from_options(options);
        self.writer.write(data, &sheet).map(|_| ())
    }
}

impl ExcelBuilder for ExcelBuilderImpl {
    fn add_content<T, I>(&mut self, data: I, write_sheet: &WriteSheet<T>) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        self.write_rows(data, write_sheet, None)
    }

    fn add_content_with_table<T, I>(
        &mut self,
        data: I,
        write_sheet: &WriteSheet<T>,
        write_table: &WriteTable,
    ) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        self.write_rows(data, write_sheet, Some(write_table))
    }

    fn merge(
        &mut self,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<()> {
        self.pending_merges
            .push(MergeRange::new(first_row, last_row, first_col, last_col));
        Ok(())
    }

    fn write_context(&self) -> &dyn WriteContext {
        &self.context
    }

    fn fill(
        &mut self,
        data: &dyn Any,
        fill_config: FillConfig,
        write_sheet: &WriteSheet<DynamicRow>,
    ) -> Result<()> {
        if !self.writer.has_template_configured() {
            return Err(fill_requires_template_error());
        }
        if is_csv_path(&self.logical_path) {
            return Err(csv_fill_unsupported_error());
        }
        if is_xls_path(&self.logical_path) {
            return Err(ExcelError::Unsupported(
                "legacy XLS template fill is not supported".to_owned(),
            ));
        }
        let executor = self.fill_executor.as_mut().ok_or_else(|| {
            ExcelError::Unsupported(
                "template fill executor is not wired; build through easyexcel::builder_from_writer"
                    .to_owned(),
            )
        })?;
        let sheet = WriteFillSheet {
            sheet_name: write_sheet.options().sheet_name.clone(),
            sheet_index: write_sheet.options().sheet_index,
        };
        executor.fill(
            data,
            WriteFillConfig {
                force_new_row: fill_config.force_new_row,
                direction: None,
            },
            sheet,
        )?;
        self.fill_session_active = true;
        Ok(())
    }

    fn finish(&mut self, on_exception: bool) -> Result<()> {
        if self.fill_session_active {
            if let Some(executor) = self.fill_executor.as_mut() {
                executor.finish(on_exception)?;
                self.writer.mark_finished();
                self.finished_via_fill = true;
                return Ok(());
            }
        }
        if on_exception {
            self.writer.finish_on_exception()
        } else {
            self.writer.finish()
        }
    }
}

fn is_csv_path(path: &std::path::Path) -> bool {
    path.extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("csv"))
}

fn is_xls_path(path: &std::path::Path) -> bool {
    path.extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xls"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use easyexcel_core::DynamicRow;
    use tempfile::tempdir;

    #[test]
    fn excel_builder_impl_delegates_add_content_and_finish() -> Result<()> {
        let directory = tempdir()?;
        let path = directory.path().join("builder-facade.xlsx");
        let sheet = WriteSheet::<DynamicRow>::new("Sheet1");
        let mut builder = ExcelBuilderImpl::from_options(&path, WriteOptions::default());
        builder.add_content(
            [DynamicRow::new({
                let mut cells = std::collections::BTreeMap::new();
                cells.insert(
                    0,
                    easyexcel_core::DynamicValue::String("alpha".to_owned()),
                );
                cells
            })],
            &sheet,
        )?;
        builder.finish(false)?;
        assert!(path.exists());
        Ok(())
    }

    #[test]
    fn excel_builder_merge_is_applied_on_next_add_content() -> Result<()> {
        let directory = tempdir()?;
        let path = directory.path().join("builder-merge.xlsx");
        let sheet = WriteSheet::<DynamicRow>::new("Sheet1");
        let mut builder = ExcelBuilderImpl::from_options(&path, WriteOptions::default());
        builder.merge(0, 0, 0, 1)?;
        builder.add_content(
            [DynamicRow::new({
                let mut cells = std::collections::BTreeMap::new();
                cells.insert(
                    0,
                    easyexcel_core::DynamicValue::String("merged".to_owned()),
                );
                cells
            })],
            &sheet,
        )?;
        builder.finish(false)?;
        assert!(path.exists());
        Ok(())
    }

    #[test]
    fn write_context_exposes_sheet_and_table_after_add_content() -> Result<()> {
        let directory = tempdir()?;
        let path = directory.path().join("builder-context.xlsx");
        let sheet = WriteSheet::<DynamicRow>::new(" Sheet1 ");
        let table = crate::ExcelWriterTableBuilder::new()
            .table_no(1)
            .need_head(false)
            .build();
        let mut builder = ExcelBuilderImpl::from_options(&path, WriteOptions::default());
        builder.add_content_with_table([], &sheet, &table)?;

        let holder = builder.write_context().current_write_holder();
        assert_eq!(
            holder.sheet_context().map(WriteSheetContext::sheet_name),
            Some("Sheet1")
        );
        assert_eq!(holder.table_no(), Some(1));
        assert!(holder.workbook_context().is_some());
        Ok(())
    }
}
