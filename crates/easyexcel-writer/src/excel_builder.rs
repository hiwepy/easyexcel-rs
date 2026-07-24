//! Mirrors Java `com.alibaba.excel.write.ExcelBuilder` and `ExcelBuilderImpl`.

use std::any::Any;
use std::path::PathBuf;

#[cfg(test)]
use easyexcel_core::Holder;
use easyexcel_core::{
    DynamicRow, ExcelError, ExcelRow, Result, WriteContext, WriteContextImpl,
    WriteContextLifecycle, WriteFillConfig, WriteFillExecutor, WriteFillSheet,
    csv_fill_unsupported_error, fill_requires_template_error, finish_write_context,
};

use crate::builder::excel_writer_table_builder::merge_table_options;
use crate::executor::excel_write_fill_executor::ExcelWriteFillExecutor;
use crate::metadata::WriteTable;
use crate::{ExcelWriter, MergeRange, WriteOptions, WriteSheet};

/// Minimal fill configuration accepted by [`ExcelBuilder::fill`].
///
/// Mirrors Java `com.alibaba.excel.write.metadata.fill.FillConfig` at the
/// builder surface. Stateful template filling remains on
/// `easyexcel_template::FillConfig`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FillConfig {
    /// Collection expansion direction. `None` is initialized as vertical.
    /// (Java `FillConfig.direction`)
    pub direction: Option<easyexcel_core::WriteDirection>,
    /// Whether collection fill forces a new row. (Java `FillConfig.forceNewRow`)
    pub force_new_row: bool,
    /// Whether generated cells inherit the template style.
    /// (Java `FillConfig.autoStyle`, default `true`)
    pub auto_style: bool,
    has_init: bool,
}

impl FillConfig {
    /// Creates Java-compatible effective defaults.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            direction: None,
            force_new_row: false,
            auto_style: true,
            has_init: false,
        }
    }

    /// Sets the collection expansion direction.
    #[must_use]
    pub const fn direction(mut self, direction: easyexcel_core::WriteDirection) -> Self {
        self.direction = Some(direction);
        self
    }

    /// Sets whether collection fill forces a new row.
    #[must_use]
    pub const fn force_new_row(mut self, force_new_row: bool) -> Self {
        self.force_new_row = force_new_row;
        self
    }

    /// Sets whether generated cells inherit the template style.
    #[must_use]
    pub const fn auto_style(mut self, auto_style: bool) -> Self {
        self.auto_style = auto_style;
        self
    }

    /// Applies Java defaults once. Rust stores effective non-null values, so
    /// initialization only records the lifecycle transition.
    pub fn init(&mut self) {
        if !self.has_init {
            self.has_init = true;
        }
    }

    /// Returns whether [`Self::init`] has run.
    #[must_use]
    pub const fn has_init(&self) -> bool {
        self.has_init
    }
}

impl Default for FillConfig {
    fn default() -> Self {
        Self::new()
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
    fn merge(&mut self, first_row: u32, last_row: u32, first_col: u16, last_col: u16)
    -> Result<()>;

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

    fn update_current_holder<T>(
        &mut self,
        options: &WriteOptions,
        table_no: Option<i32>,
    ) -> Result<()>
    where
        T: ExcelRow,
    {
        self.context.set_sheet_context(&options.sheet_name);
        self.context.set_table_no(table_no);
        self.context
            .set_current_holder_state(crate::resolved_write_context_holder_state::<T>(
                options, table_no,
            )?);
        Ok(())
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
        self.update_current_holder::<T>(&options, write_table.map(WriteTable::table_no))?;
        let sheet = WriteSheet::from_options(options);
        self.writer.write(data, &sheet).map(|_| ())
    }

    fn finish_resources(&mut self, on_exception: bool) -> Result<()> {
        if self.fill_session_active {
            if let Some(delegate) = self.fill_executor.as_mut() {
                let mut executor =
                    ExcelWriteFillExecutor::with_delegate(&self.context, delegate.as_mut());
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

impl WriteContext for ExcelBuilderImpl {
    fn current_write_holder(&self) -> &dyn easyexcel_core::WriteContextHolder {
        self.context.current_write_holder()
    }
}

impl WriteContextLifecycle for ExcelBuilderImpl {
    fn finish_context(&mut self, on_exception: bool) -> Result<()> {
        self.finish_resources(on_exception)
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
        mut fill_config: FillConfig,
        write_sheet: &WriteSheet<DynamicRow>,
    ) -> Result<()> {
        fill_config.init();
        if !self.writer.has_template_configured() {
            return Err(fill_requires_template_error());
        }
        if self.writer.is_csv() {
            return Err(csv_fill_unsupported_error());
        }
        if self.writer.is_xls() {
            return Err(ExcelError::Unsupported(
                "legacy XLS template fill is not supported".to_owned(),
            ));
        }
        let mut holder_options = write_sheet.options().clone();
        holder_options.sheet_name = if holder_options.auto_trim {
            holder_options.sheet_name.trim().to_owned()
        } else {
            holder_options.sheet_name.clone()
        };
        self.update_current_holder::<DynamicRow>(&holder_options, None)?;
        let delegate = self.fill_executor.as_mut().ok_or_else(|| {
            ExcelError::Unsupported(
                "template fill executor is not wired; build through easyexcel::builder_from_writer"
                    .to_owned(),
            )
        })?;
        let sheet = WriteFillSheet {
            sheet_name: write_sheet.options().sheet_name.clone(),
            sheet_index: write_sheet.options().sheet_index,
        };
        let mut executor = ExcelWriteFillExecutor::with_delegate(&self.context, delegate.as_mut());
        executor.fill(
            data,
            WriteFillConfig {
                force_new_row: fill_config.force_new_row,
                direction: fill_config.direction,
                auto_style: fill_config.auto_style,
            },
            sheet,
        )?;
        self.fill_session_active = true;
        Ok(())
    }

    fn finish(&mut self, on_exception: bool) -> Result<()> {
        finish_write_context(self, on_exception)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use easyexcel_core::{
        CellValue, DynamicRow, ExcelColumn, ExcelWriteMetadata, RowData, WriteDirection,
        WriteSheetContext,
    };
    use tempfile::tempdir;

    struct ContextRow;

    #[derive(Default)]
    struct ContextFillExecutor;

    impl WriteFillExecutor for ContextFillExecutor {
        fn fill(
            &mut self,
            _data: &dyn Any,
            _fill_config: WriteFillConfig,
            _sheet: WriteFillSheet,
        ) -> Result<()> {
            Ok(())
        }

        fn finish(&mut self, _on_exception: bool) -> Result<()> {
            Ok(())
        }
    }

    impl ExcelRow for ContextRow {
        fn schema() -> &'static [ExcelColumn] {
            static SCHEMA: [ExcelColumn; 3] = [
                ExcelColumn::new("a", "A", Some(0), 0, None),
                ExcelColumn::new("b", "B", Some(1), 0, None),
                ExcelColumn::new("c", "C", Some(2), 0, None),
            ];
            &SCHEMA
        }

        fn write_metadata() -> &'static ExcelWriteMetadata {
            static METADATA: ExcelWriteMetadata = ExcelWriteMetadata::new().head_row_height(28);
            &METADATA
        }

        fn from_row(_row: &RowData) -> Result<Self> {
            Ok(Self)
        }

        fn to_row(&self) -> Result<Vec<CellValue>> {
            Ok(vec![
                CellValue::String("a".to_owned()),
                CellValue::String("b".to_owned()),
                CellValue::String("c".to_owned()),
            ])
        }
    }

    #[test]
    fn fill_config_initializes_java_defaults_and_preserves_overrides() {
        let mut defaults = FillConfig::new();
        assert_eq!(defaults.direction, None);
        assert!(!defaults.force_new_row);
        assert!(defaults.auto_style);
        assert!(!defaults.has_init());
        defaults.init();
        defaults.init();
        assert!(defaults.has_init());

        let configured = FillConfig::new()
            .direction(WriteDirection::Horizontal)
            .force_new_row(true)
            .auto_style(false);
        assert_eq!(configured.direction, Some(WriteDirection::Horizontal));
        assert!(configured.force_new_row);
        assert!(!configured.auto_style);
    }

    #[test]
    fn fill_uses_explicit_excel_type_instead_of_path_extension() {
        let mut builder = ExcelBuilderImpl::from_options(
            "logical.xlsx",
            WriteOptions {
                excel_type: Some(easyexcel_core::support::ExcelTypeEnum::Csv),
                template_bytes: Some(vec![1]),
                ..WriteOptions::default()
            },
        );
        let sheet = WriteSheet::<DynamicRow>::new("Sheet1");
        let error = builder
            .fill(&DynamicRow::default(), FillConfig::new(), &sheet)
            .expect_err("explicit CSV type must reject template fill");
        assert_eq!(
            error.to_string(),
            "unsupported operation: csv does not support filling data."
        );
    }

    #[test]
    fn excel_builder_impl_delegates_add_content_and_finish() -> Result<()> {
        let directory = tempdir()?;
        let path = directory.path().join("builder-facade.xlsx");
        let sheet = WriteSheet::<DynamicRow>::new("Sheet1");
        let mut builder = ExcelBuilderImpl::from_options(&path, WriteOptions::default());
        builder.add_content(
            [DynamicRow::new({
                let mut cells = std::collections::BTreeMap::new();
                cells.insert(0, easyexcel_core::DynamicValue::String("alpha".to_owned()));
                cells
            })],
            &sheet,
        )?;
        finish_write_context(&mut builder, false)?;
        finish_write_context(&mut builder, false)?;
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
                cells.insert(0, easyexcel_core::DynamicValue::String("merged".to_owned()));
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

    #[test]
    fn live_current_write_holder_tracks_resolved_sheet_and_table_state() -> Result<()> {
        let directory = tempdir()?;
        let path = directory.path().join("live-holder-context.xlsx");
        let mut builder = ExcelBuilderImpl::from_options(&path, WriteOptions::default());
        let sheet = WriteSheet::<ContextRow>::from_options(WriteOptions {
            sheet_name: " Typed ".to_owned(),
            include_column_indexes: Some(vec![2, 0]),
            order_by_include_column: true,
            relative_head_row_index: 3,
            automatic_merge_head: false,
            dynamic_head: Some(vec![
                vec!["Group".to_owned(), "A*".to_owned()],
                vec!["Group".to_owned(), "B*".to_owned()],
                vec!["Group".to_owned(), "C*".to_owned()],
            ]),
            ..WriteOptions::default()
        });
        builder.add_content([], &sheet)?;

        let sheet_holder = builder.write_context().current_write_holder();
        assert_eq!(sheet_holder.holder_type(), Holder::Sheet);
        assert_eq!(
            sheet_holder
                .sheet_context()
                .map(WriteSheetContext::sheet_name),
            Some("Typed")
        );
        assert!(sheet_holder.need_head());
        assert!(!sheet_holder.automatic_merge_head());
        assert_eq!(sheet_holder.relative_head_row_index(), 3);
        assert!(sheet_holder.order_by_include_column());
        assert_eq!(
            sheet_holder.include_column_indexes(),
            Some([2, 0].as_slice())
        );
        assert_eq!(
            sheet_holder
                .excel_write_head_property()
                .head_row_height_property()
                .map(easyexcel_core::metadata::RowHeightProperty::height),
            Some(28)
        );
        assert_eq!(
            sheet_holder
                .excel_write_head_property()
                .head_map()
                .values()
                .map(|head| (
                    head.column_index(),
                    head.field_name().map(str::to_owned),
                    head.head_name_list().to_vec(),
                ))
                .collect::<Vec<_>>(),
            vec![
                (
                    Some(0),
                    Some("c".to_owned()),
                    vec!["Group".to_owned(), "C*".to_owned()]
                ),
                (
                    Some(1),
                    Some("a".to_owned()),
                    vec!["Group".to_owned(), "A*".to_owned()]
                ),
            ]
        );

        let table = crate::ExcelWriterTableBuilder::new()
            .table_no(7)
            .need_head(false)
            .include_column_field_names(["b"])
            .build();
        builder.add_content_with_table([], &sheet, &table)?;
        let table_holder = builder.write_context().current_write_holder();
        assert_eq!(table_holder.holder_type(), Holder::Table);
        assert_eq!(table_holder.table_no(), Some(7));
        assert!(!table_holder.need_head());
        assert_eq!(
            table_holder.include_column_field_names(),
            Some(["b".to_owned()].as_slice())
        );
        assert_eq!(
            table_holder.include_column_indexes(),
            Some([2, 0].as_slice())
        );
        assert_eq!(
            table_holder
                .excel_write_head_property()
                .head_map()
                .values()
                .map(|head| head.field_name())
                .collect::<Vec<_>>(),
            vec![Some("b"), Some("c"), Some("a")]
        );
        builder.finish(false)?;
        assert!(path.exists());
        Ok(())
    }

    #[test]
    fn template_fill_updates_the_same_live_current_holder() -> Result<()> {
        let mut builder = ExcelBuilderImpl::from_options(
            "fill-context.xlsx",
            WriteOptions {
                template_bytes: Some(vec![1]),
                ..WriteOptions::default()
            },
        );
        builder.set_fill_executor(Box::new(ContextFillExecutor));
        let sheet = WriteSheet::<DynamicRow>::from_options(WriteOptions {
            sheet_name: " Fill ".to_owned(),
            need_head: false,
            relative_head_row_index: 4,
            automatic_merge_head: false,
            dynamic_head: Some(vec![vec!["填充列".to_owned()]]),
            ..WriteOptions::default()
        });
        builder.fill(&DynamicRow::default(), FillConfig::new(), &sheet)?;

        let holder = builder.write_context().current_write_holder();
        assert_eq!(holder.holder_type(), Holder::Sheet);
        assert_eq!(
            holder.sheet_context().map(WriteSheetContext::sheet_name),
            Some("Fill")
        );
        assert!(!holder.need_head());
        assert!(!holder.automatic_merge_head());
        assert_eq!(holder.relative_head_row_index(), 4);
        assert_eq!(
            holder
                .excel_write_head_property()
                .head_map()
                .values()
                .flat_map(|head| head.head_name_list())
                .map(String::as_str)
                .collect::<Vec<_>>(),
            vec!["填充列"]
        );
        Ok(())
    }
}
