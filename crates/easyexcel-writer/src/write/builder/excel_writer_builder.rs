//! Rust implementation of Java
//! `com.alibaba.excel.write.builder.ExcelWriterBuilder`.

use std::path::PathBuf;

use easyexcel_core::support::ExcelTypeEnum;
use easyexcel_core::{CsvCharset, ExcelError, Result, WriteHandler};

use crate::metadata::write_workbook::WriteWorkbook;
use crate::write::builder::excel_writer_sheet_builder::ExcelWriterSheetBuilder;
use crate::{ExcelOutputStream, ExcelWriter};

/// Java-compatible workbook writer builder backed by the real Rust writer.
///
/// Unlike the former documentation-only placeholder, every supported option
/// is stored on [`WriteWorkbook`] and is passed into [`ExcelWriter`] by
/// [`Self::build`].
pub struct ExcelWriterBuilder {
    write_workbook: WriteWorkbook,
    handlers: Vec<Box<dyn WriteHandler>>,
}

impl ExcelWriterBuilder {
    /// Creates an empty builder. (Java `ExcelWriterBuilder()`)
    #[must_use]
    pub fn new() -> Self {
        Self {
            write_workbook: WriteWorkbook::new(),
            handlers: Vec::new(),
        }
    }

    /// Sets the final output file. (Java `file(File/String)`)
    #[must_use]
    pub fn file(mut self, file: impl Into<PathBuf>) -> Self {
        self.write_workbook.set_file(file);
        self
    }

    /// Sets the requested workbook type. (Java `excelType(ExcelTypeEnum)`)
    #[must_use]
    pub fn excel_type(mut self, excel_type: ExcelTypeEnum) -> Self {
        self.write_workbook.set_excel_type(excel_type);
        self
    }

    /// Enables or disables Java's default bold header style.
    #[must_use]
    pub fn use_default_style(mut self, enabled: bool) -> Self {
        self.write_workbook.options.use_default_style = enabled;
        self.write_workbook.options.head_style = if enabled {
            crate::CellStyle::new().bold(true)
        } else {
            crate::CellStyle::new()
        };
        self
    }

    /// Controls whether an owned output stream closes on finish.
    #[must_use]
    pub fn auto_close_stream(mut self, enabled: bool) -> Self {
        self.write_workbook.set_auto_close_stream(enabled);
        self
    }

    /// Encrypts XLSX output.
    #[must_use]
    pub fn password(mut self, password: impl Into<String>) -> Self {
        self.write_workbook.set_password(password);
        self
    }

    /// Selects in-memory instead of constant-memory writing.
    #[must_use]
    pub fn in_memory(mut self, enabled: bool) -> Self {
        self.write_workbook.set_in_memory(enabled);
        self
    }

    /// Controls whether partial output is emitted on an exception.
    #[must_use]
    pub fn write_excel_on_exception(mut self, enabled: bool) -> Self {
        self.write_workbook.set_write_excel_on_exception(enabled);
        self
    }

    /// Sets CSV output encoding.
    #[must_use]
    pub fn charset(mut self, charset: impl Into<CsvCharset>) -> Self {
        self.write_workbook.options.charset = charset.into();
        self
    }

    /// Controls whether CSV starts with a byte-order mark.
    #[must_use]
    pub fn with_bom(mut self, enabled: bool) -> Self {
        self.write_workbook.set_with_bom(enabled);
        self
    }

    /// Sets a template file. (Java `withTemplate(File/String)`)
    #[must_use]
    pub fn with_template(mut self, template_file: impl Into<PathBuf>) -> Self {
        self.write_workbook.set_template_file(template_file);
        self
    }

    /// Sets a buffered template stream. (Java `withTemplate(InputStream)`)
    #[must_use]
    pub fn with_template_bytes(mut self, template_bytes: impl Into<Vec<u8>>) -> Self {
        self.write_workbook.set_template_bytes(template_bytes);
        self
    }

    /// Sets the number of rows before the header.
    #[must_use]
    pub fn relative_head_row_index(mut self, index: i32) -> Self {
        self.write_workbook.options.relative_head_row_index = index;
        self
    }

    /// Controls header output.
    #[must_use]
    pub fn need_head(mut self, enabled: bool) -> Self {
        self.write_workbook.options.need_head = enabled;
        self
    }

    /// Controls automatic merging of equal multi-level header cells.
    #[must_use]
    pub fn automatic_merge_head(mut self, enabled: bool) -> Self {
        self.write_workbook.options.automatic_merge_head = enabled;
        self
    }

    /// Includes only the supplied physical columns.
    #[must_use]
    pub fn include_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        self.write_workbook.options.include_column_indexes = Some(indexes.into_iter().collect());
        self
    }

    /// Includes only the supplied Rust field names.
    #[must_use]
    pub fn include_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        self.write_workbook.options.include_column_field_names =
            Some(names.into_iter().map(Into::into).collect());
        self
    }

    /// Excludes physical columns.
    #[must_use]
    pub fn exclude_column_indexes(mut self, indexes: impl IntoIterator<Item = usize>) -> Self {
        self.write_workbook.options.exclude_column_indexes = indexes.into_iter().collect();
        self
    }

    /// Excludes Rust field names.
    #[must_use]
    pub fn exclude_column_field_names<S>(mut self, names: impl IntoIterator<Item = S>) -> Self
    where
        S: Into<String>,
    {
        self.write_workbook.options.exclude_column_field_names =
            names.into_iter().map(Into::into).collect();
        self
    }

    /// Orders output by the include-list order.
    #[must_use]
    pub fn order_by_include_column(mut self, enabled: bool) -> Self {
        self.write_workbook.options.order_by_include_column = enabled;
        self
    }

    /// Appends a write handler in registration order.
    #[must_use]
    pub fn register_write_handler(mut self, handler: impl WriteHandler + 'static) -> Self {
        self.handlers.push(Box::new(handler));
        self
    }

    /// Selects an owned output stream instead of a file.
    ///
    /// This is Rust's typed equivalent of Java `file(OutputStream)`. Configure
    /// workbook options before this call, then call `build` or `sheet` on the
    /// returned stream builder.
    #[must_use]
    pub fn output_stream<W>(self, output: ExcelOutputStream<W>) -> ExcelWriterOutputStreamBuilder<W>
    where
        W: std::io::Write + Send + 'static,
    {
        ExcelWriterOutputStreamBuilder {
            builder: self,
            output,
        }
    }

    /// Returns the accumulated Java-style metadata.
    #[must_use]
    pub const fn parameter(&self) -> &WriteWorkbook {
        &self.write_workbook
    }

    /// Builds a stateful writer. (Java `build()`)
    pub fn build(self) -> Result<ExcelWriter> {
        let path = self.write_workbook.output_file.ok_or_else(|| {
            ExcelError::Format("ExcelWriterBuilder.file must be set before build()".to_owned())
        })?;
        Ok(ExcelWriter::with_handlers_and_options(
            path,
            self.handlers,
            self.write_workbook.options,
        ))
    }

    /// Builds a writer-bound default sheet.
    pub fn sheet(self) -> Result<ExcelWriterSheetBuilder> {
        let inherited_options = self.write_workbook.options.clone();
        Ok(ExcelWriterSheetBuilder::with_excel_writer_and_options(
            self.build()?,
            inherited_options,
        ))
    }

    /// Builds a writer-bound sheet selected by number.
    pub fn sheet_no(self, sheet_no: i32) -> Result<ExcelWriterSheetBuilder> {
        Ok(self.sheet()?.sheet_no(sheet_no))
    }

    /// Builds a writer-bound sheet selected by name.
    pub fn sheet_name(self, sheet_name: impl Into<String>) -> Result<ExcelWriterSheetBuilder> {
        Ok(self.sheet()?.sheet_name(sheet_name))
    }

    /// Builds a writer-bound sheet selected by number and name.
    pub fn sheet_with(
        self,
        sheet_no: i32,
        sheet_name: impl Into<String>,
    ) -> Result<ExcelWriterSheetBuilder> {
        Ok(self.sheet()?.sheet_no(sheet_no).sheet_name(sheet_name))
    }
}

impl Default for ExcelWriterBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Java-compatible writer builder whose destination is an output stream.
pub struct ExcelWriterOutputStreamBuilder<W> {
    builder: ExcelWriterBuilder,
    output: ExcelOutputStream<W>,
}

impl<W> ExcelWriterOutputStreamBuilder<W>
where
    W: std::io::Write + Send + 'static,
{
    /// Builds a stateful writer backed by the configured stream.
    #[must_use]
    pub fn build(self) -> ExcelWriter {
        let logical_path = self
            .builder
            .write_workbook
            .output_file
            .unwrap_or_else(|| PathBuf::from("easyexcel.xlsx"));
        ExcelWriter::with_output_stream(
            logical_path,
            self.output,
            self.builder.handlers,
            self.builder.write_workbook.options,
        )
    }

    /// Builds a writer-bound default sheet.
    #[must_use]
    pub fn sheet(self) -> ExcelWriterSheetBuilder {
        let inherited_options = self.builder.write_workbook.options.clone();
        ExcelWriterSheetBuilder::with_excel_writer_and_options(self.build(), inherited_options)
    }

    /// Builds a writer-bound sheet selected by number.
    #[must_use]
    pub fn sheet_no(self, sheet_no: i32) -> ExcelWriterSheetBuilder {
        self.sheet().sheet_no(sheet_no)
    }

    /// Builds a writer-bound sheet selected by name.
    #[must_use]
    pub fn sheet_name(self, sheet_name: impl Into<String>) -> ExcelWriterSheetBuilder {
        self.sheet().sheet_name(sheet_name)
    }

    /// Builds a writer-bound sheet selected by number and name.
    #[must_use]
    pub fn sheet_with(
        self,
        sheet_no: i32,
        sheet_name: impl Into<String>,
    ) -> ExcelWriterSheetBuilder {
        self.sheet().sheet_no(sheet_no).sheet_name(sheet_name)
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;
    use std::io::Read as _;
    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::sync::{Arc, Mutex};

    use calamine::{DataType, Reader, Xlsx, open_workbook};
    use easyexcel_core::event::NotRepeatExecutor;
    use easyexcel_core::{
        CellValue, ExcelColumn, ExcelRow, RowData, WriteCellContext, WriteRowContext,
        WriteSheetContext, WriteWorkbookContext,
    };
    use tempfile::tempdir;

    use super::*;

    struct SimpleRow(&'static str);

    impl ExcelRow for SimpleRow {
        fn schema() -> &'static [ExcelColumn] {
            const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Value", Some(0), 0, None)];
            COLUMNS
        }

        fn from_row(_row: &RowData) -> Result<Self> {
            Ok(Self(""))
        }

        fn to_row(&self) -> Result<Vec<CellValue>> {
            Ok(vec![CellValue::String(self.0.to_owned())])
        }
    }

    struct TwoColumnRow(&'static str, &'static str);

    impl ExcelRow for TwoColumnRow {
        fn schema() -> &'static [ExcelColumn] {
            const COLUMNS: &[ExcelColumn] = &[
                ExcelColumn::new("first", "First", Some(0), 0, None),
                ExcelColumn::new("second", "Second", Some(1), 1, None),
            ];
            COLUMNS
        }

        fn from_row(_row: &RowData) -> Result<Self> {
            Ok(Self("", ""))
        }

        fn to_row(&self) -> Result<Vec<CellValue>> {
            Ok(vec![
                CellValue::String(self.0.to_owned()),
                CellValue::String(self.1.to_owned()),
            ])
        }
    }

    struct WorkbookProbe(Arc<AtomicUsize>);

    impl WriteHandler for WorkbookProbe {
        fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.0.fetch_add(1, Ordering::SeqCst);
            Ok(())
        }
    }

    struct UniqueWorkbookProbe {
        calls: Arc<AtomicUsize>,
        order: i32,
        unique_value: &'static str,
    }

    impl NotRepeatExecutor for UniqueWorkbookProbe {
        fn unique_value(&self) -> &str {
            self.unique_value
        }
    }

    impl WriteHandler for UniqueWorkbookProbe {
        fn order(&self) -> i32 {
            self.order
        }

        fn as_not_repeat_executor(&self) -> Option<&dyn NotRepeatExecutor> {
            Some(self)
        }

        fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.calls.fetch_add(1, Ordering::SeqCst);
            Ok(())
        }
    }

    struct ExactLifecycleProbe(Arc<Mutex<Vec<&'static str>>>);

    impl ExactLifecycleProbe {
        fn record(&self, event: &'static str) {
            self.0.lock().expect("event log mutex poisoned").push(event);
        }
    }

    impl WriteHandler for ExactLifecycleProbe {
        fn before_workbook_create(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.record("before_workbook_create");
            Ok(())
        }

        fn after_workbook_create(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.record("after_workbook_create");
            Ok(())
        }

        fn after_workbook_dispose(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.record("after_workbook_dispose");
            Ok(())
        }

        fn before_sheet_create(&mut self, _context: &WriteSheetContext) -> Result<()> {
            self.record("before_sheet_create");
            Ok(())
        }

        fn after_sheet_create(&mut self, _context: &WriteSheetContext) -> Result<()> {
            self.record("after_sheet_create");
            Ok(())
        }

        fn after_sheet_dispose(&mut self, _context: &WriteSheetContext) -> Result<()> {
            self.record("after_sheet_dispose");
            Ok(())
        }

        fn before_row_create(&mut self, _context: &WriteRowContext) -> Result<()> {
            self.record("before_row_create");
            Ok(())
        }

        fn after_row_create(&mut self, _context: &WriteRowContext) -> Result<()> {
            self.record("after_row_create");
            Ok(())
        }

        fn after_row_dispose(&mut self, _context: &WriteRowContext) -> Result<()> {
            self.record("after_row_dispose");
            Ok(())
        }

        fn before_cell_create(&mut self, _context: &mut WriteCellContext) -> Result<()> {
            self.record("before_cell_create");
            Ok(())
        }

        fn after_cell_create(&mut self, _context: &WriteCellContext) -> Result<()> {
            self.record("after_cell_create");
            Ok(())
        }

        fn after_cell_data_converted(&mut self, _context: &WriteCellContext) -> Result<()> {
            self.record("after_cell_data_converted");
            Ok(())
        }

        fn after_cell_dispose(&mut self, _context: &WriteCellContext) -> Result<()> {
            self.record("after_cell_dispose");
            Ok(())
        }
    }

    fn zip_entry(path: &std::path::Path, name: &str) -> Result<String> {
        let file = std::fs::File::open(path)?;
        let mut archive =
            zip::ZipArchive::new(file).map_err(|error| ExcelError::Format(error.to_string()))?;
        let mut entry = archive
            .by_name(name)
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        let mut text = String::new();
        entry.read_to_string(&mut text)?;
        Ok(text)
    }

    #[test]
    fn compatibility_builder_writes_xlsx_and_runs_registered_handlers() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("users.xlsx");
        let calls = Arc::new(AtomicUsize::new(0));

        ExcelWriterBuilder::new()
            .file(&output)
            .need_head(false)
            .register_write_handler(WorkbookProbe(Arc::clone(&calls)))
            .sheet_name("Users")?
            .register_write_handler(WorkbookProbe(Arc::clone(&calls)))
            .do_write(vec![SimpleRow("alice")])?;

        assert_eq!(calls.load(Ordering::SeqCst), 2);
        let mut workbook: Xlsx<_> = open_workbook(&output)
            .map_err(|error: calamine::XlsxError| ExcelError::Format(error.to_string()))?;
        let range = workbook
            .worksheet_range("Users")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        assert_eq!(
            range.get_value((0, 0)).and_then(|cell| cell.get_string()),
            Some("alice")
        );
        Ok(())
    }

    #[test]
    fn exact_java_handler_lifecycle_runs_in_creation_and_dispose_order() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("exact-handler-lifecycle.xlsx");
        let events = Arc::new(Mutex::new(Vec::new()));

        ExcelWriterBuilder::new()
            .file(&output)
            .need_head(false)
            .register_write_handler(ExactLifecycleProbe(Arc::clone(&events)))
            .sheet_name("Users")?
            .do_write(vec![SimpleRow("alice")])?;

        assert_eq!(
            *events.lock().expect("event log mutex poisoned"),
            vec![
                "before_workbook_create",
                "after_workbook_create",
                "before_sheet_create",
                "after_sheet_create",
                "before_row_create",
                "after_row_create",
                "before_cell_create",
                "after_cell_create",
                "after_cell_data_converted",
                "after_cell_dispose",
                "after_row_dispose",
                "after_sheet_dispose",
                "after_workbook_dispose",
            ]
        );
        Ok(())
    }

    #[test]
    fn duplicate_handlers_execute_only_the_lowest_order_instance() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("deduplicated-handlers.xlsx");
        let first = Arc::new(AtomicUsize::new(0));
        let duplicate = Arc::new(AtomicUsize::new(0));
        let repeatable = Arc::new(AtomicUsize::new(0));

        ExcelWriterBuilder::new()
            .file(&output)
            .register_write_handler(UniqueWorkbookProbe {
                calls: Arc::clone(&duplicate),
                order: 20,
                unique_value: "workbook-probe",
            })
            .register_write_handler(WorkbookProbe(Arc::clone(&repeatable)))
            .register_write_handler(UniqueWorkbookProbe {
                calls: Arc::clone(&first),
                order: -20,
                unique_value: "workbook-probe",
            })
            .register_write_handler(WorkbookProbe(Arc::clone(&repeatable)))
            .sheet_name("Users")?
            .do_write(vec![SimpleRow("alice")])?;

        assert_eq!(first.load(Ordering::SeqCst), 1);
        assert_eq!(duplicate.load(Ordering::SeqCst), 0);
        assert_eq!(repeatable.load(Ordering::SeqCst), 2);
        Ok(())
    }

    #[test]
    fn sheet_own_workbook_callback_is_supplementary_to_initialized_parent() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("sheet-handler-precedence.xlsx");
        let workbook_calls = Arc::new(AtomicUsize::new(0));
        let sheet_calls = Arc::new(AtomicUsize::new(0));

        ExcelWriterBuilder::new()
            .file(&output)
            .register_write_handler(UniqueWorkbookProbe {
                calls: Arc::clone(&workbook_calls),
                order: 0,
                unique_value: "same-handler",
            })
            .sheet_name("Users")?
            .register_write_handler(UniqueWorkbookProbe {
                calls: Arc::clone(&sheet_calls),
                order: 0,
                unique_value: "same-handler",
            })
            .do_write(vec![SimpleRow("alice")])?;

        assert_eq!(sheet_calls.load(Ordering::SeqCst), 1);
        assert_eq!(workbook_calls.load(Ordering::SeqCst), 1);
        Ok(())
    }

    #[test]
    fn table_and_sheet_own_workbook_callbacks_are_each_supplementary() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("table-handler-precedence.xlsx");
        let workbook_calls = Arc::new(AtomicUsize::new(0));
        let sheet_calls = Arc::new(AtomicUsize::new(0));
        let table_calls = Arc::new(AtomicUsize::new(0));

        ExcelWriterBuilder::new()
            .file(&output)
            .register_write_handler(UniqueWorkbookProbe {
                calls: Arc::clone(&workbook_calls),
                order: 0,
                unique_value: "same-handler",
            })
            .sheet_name("Users")?
            .register_write_handler(UniqueWorkbookProbe {
                calls: Arc::clone(&sheet_calls),
                order: 0,
                unique_value: "same-handler",
            })
            .table()
            .register_write_handler(Box::new(UniqueWorkbookProbe {
                calls: Arc::clone(&table_calls),
                order: 0,
                unique_value: "same-handler",
            }))
            .do_write(vec![SimpleRow("alice")])?;

        assert_eq!(table_calls.load(Ordering::SeqCst), 1);
        assert_eq!(sheet_calls.load(Ordering::SeqCst), 1);
        assert_eq!(workbook_calls.load(Ordering::SeqCst), 1);
        Ok(())
    }

    #[test]
    fn sheet_explicit_need_head_overrides_inherited_workbook_value() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("sheet-override.xlsx");

        ExcelWriterBuilder::new()
            .file(&output)
            .need_head(false)
            .sheet_name("Users")?
            .need_head(true)
            .do_write(vec![SimpleRow("alice")])?;

        let mut workbook: Xlsx<_> = open_workbook(&output)
            .map_err(|error: calamine::XlsxError| ExcelError::Format(error.to_string()))?;
        let range = workbook
            .worksheet_range("Users")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        assert_eq!(
            range.get_value((0, 0)).and_then(|cell| cell.get_string()),
            Some("Value")
        );
        assert_eq!(
            range.get_value((1, 0)).and_then(|cell| cell.get_string()),
            Some("alice")
        );
        Ok(())
    }

    #[test]
    fn sheet_default_style_inherits_and_can_override_workbook_value() -> Result<()> {
        let directory = tempdir()?;
        let inherited_output = directory.path().join("style-inherited.xlsx");
        let overridden_output = directory.path().join("style-overridden.xlsx");

        ExcelWriterBuilder::new()
            .file(&inherited_output)
            .use_default_style(false)
            .sheet_name("Users")?
            .do_write(vec![SimpleRow("alice")])?;
        ExcelWriterBuilder::new()
            .file(&overridden_output)
            .use_default_style(false)
            .sheet_name("Users")?
            .use_default_style(true)
            .do_write(vec![SimpleRow("alice")])?;

        let inherited_styles = zip_entry(&inherited_output, "xl/styles.xml")?;
        let overridden_styles = zip_entry(&overridden_output, "xl/styles.xml")?;
        assert!(!inherited_styles.contains("<b/>"));
        assert!(overridden_styles.contains("<b/>"));
        Ok(())
    }

    #[test]
    fn explicit_excel_type_overrides_the_output_file_extension() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("users.data");

        ExcelWriterBuilder::new()
            .file(&output)
            .excel_type(ExcelTypeEnum::Csv)
            .with_bom(false)
            .sheet()?
            .do_write(vec![SimpleRow("alice")])?;

        let csv = std::fs::read_to_string(output)?;
        assert_eq!(csv, "Value\nalice\n");
        Ok(())
    }

    #[test]
    fn output_stream_builder_writes_real_xlsx_without_creating_a_file() -> Result<()> {
        let directory = tempdir()?;
        let logical_path = directory.path().join("stream.xlsx");
        let output = ExcelOutputStream::new(Cursor::new(Vec::<u8>::new()));
        let inspection = output.clone();

        ExcelWriterBuilder::new()
            .file(&logical_path)
            .auto_close_stream(false)
            .output_stream(output)
            .sheet_name("Users")
            .need_head(false)
            .do_write(vec![SimpleRow("alice")])?;

        let bytes = inspection
            .with_inner(|cursor| cursor.get_ref().clone())
            .expect("auto_close_stream(false) must keep the stream open");
        let mut archive = zip::ZipArchive::new(Cursor::new(bytes))
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        assert!(archive.by_name("[Content_Types].xml").is_ok());
        assert!(!logical_path.exists());
        Ok(())
    }

    #[test]
    fn compatibility_sheet_table_chain_writes_and_finishes() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("table-chain.xlsx");

        ExcelWriterBuilder::new()
            .file(&output)
            .sheet_name("Users")?
            .need_head(false)
            .table_no(2)
            .do_write(vec![SimpleRow("alice")])?;

        let mut workbook: Xlsx<_> = open_workbook(&output)
            .map_err(|error: calamine::XlsxError| ExcelError::Format(error.to_string()))?;
        let range = workbook
            .worksheet_range("Users")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        assert_eq!(
            range.get_value((0, 0)).and_then(|cell| cell.get_string()),
            Some("alice")
        );
        Ok(())
    }

    #[test]
    fn table_explicit_need_head_overrides_inherited_sheet_value() -> Result<()> {
        let directory = tempdir()?;
        let output = directory.path().join("table-override.xlsx");

        ExcelWriterBuilder::new()
            .file(&output)
            .need_head(false)
            .sheet_name("Users")?
            .table_no(2)
            .need_head(true)
            .do_write(vec![SimpleRow("alice")])?;

        let mut workbook: Xlsx<_> = open_workbook(&output)
            .map_err(|error: calamine::XlsxError| ExcelError::Format(error.to_string()))?;
        let range = workbook
            .worksheet_range("Users")
            .map_err(|error| ExcelError::Format(error.to_string()))?;
        assert_eq!(
            range.get_value((0, 0)).and_then(|cell| cell.get_string()),
            Some("Value")
        );
        assert_eq!(
            range.get_value((1, 0)).and_then(|cell| cell.get_string()),
            Some("alice")
        );
        Ok(())
    }

    #[test]
    fn table_column_selection_inherits_and_can_override_parent() -> Result<()> {
        let directory = tempdir()?;
        let inherited_output = directory.path().join("table-inherited-columns.xlsx");
        let overridden_output = directory.path().join("table-overridden-columns.xlsx");

        ExcelWriterBuilder::new()
            .file(&inherited_output)
            .need_head(false)
            .include_column_indexes([1])
            .sheet_name("Users")?
            .table_no(0)
            .do_write(vec![TwoColumnRow("A", "B")])?;
        ExcelWriterBuilder::new()
            .file(&overridden_output)
            .need_head(false)
            .include_column_indexes([1])
            .sheet_name("Users")?
            .table_no(0)
            .include_column_indexes([0])
            .do_write(vec![TwoColumnRow("A", "B")])?;

        for (path, expected) in [(&inherited_output, "B"), (&overridden_output, "A")] {
            let mut workbook: Xlsx<_> = open_workbook(path)
                .map_err(|error: calamine::XlsxError| ExcelError::Format(error.to_string()))?;
            let range = workbook
                .worksheet_range("Users")
                .map_err(|error| ExcelError::Format(error.to_string()))?;
            let values = range
                .rows()
                .flat_map(|row| row.iter())
                .filter_map(|cell| cell.get_string())
                .collect::<Vec<_>>();
            assert_eq!(values, vec![expected]);
        }
        Ok(())
    }

    #[test]
    fn write_workbook_file_and_template_setters_store_real_paths() {
        let mut workbook = WriteWorkbook::new();
        workbook.set_file("result.xlsx");
        workbook.set_template_file("template.xlsx");

        assert_eq!(workbook.file(), Some(std::path::Path::new("result.xlsx")));
        assert_eq!(
            workbook.template_file(),
            Some(std::path::Path::new("template.xlsx"))
        );

        let workbook = WriteWorkbook::from(crate::WriteOptions {
            excel_type: Some(ExcelTypeEnum::Csv),
            ..crate::WriteOptions::default()
        });
        assert_eq!(workbook.excel_type(), ExcelTypeEnum::Csv);
    }
}
