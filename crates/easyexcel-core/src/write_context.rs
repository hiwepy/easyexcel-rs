//! Mirrors Java `com.alibaba.excel.context.WriteContext` (interface).

use std::path::{Path, PathBuf};

use crate::ConverterRegistry;
use crate::ExcelWriteHeadProperty;
use crate::Holder;
use crate::WriteSheetContext;
use crate::WriteWorkbookContext;
use crate::excel_error::ExcelError;

/// Mirrors Java `WriteContext` (110-line interface).
///
/// Java exposes a single `currentWriteHolder()` accessor plus the
/// `finish(boolean onException)` lifecycle. Rust collapses the
/// interface to a marker struct so dependents can take a `&WriteContext`
/// reference without depending on `rust_xlsxwriter` types.
pub trait WriteContext {
    /// Returns the active write holder. (Java `WriteContext.currentWriteHolder()`)
    fn current_write_holder(&self) -> &dyn WriteContextHolder;
}

/// Resource-owning lifecycle capability for a [`WriteContext`].
///
/// Java's `WriteContextImpl` owns the POI workbook and can therefore implement
/// `finish(boolean)` directly. Rust keeps backend resources in the writer crate,
/// so only the resource-owning adapter implements this trait. Metadata-only
/// contexts deliberately do not pretend that they can persist or close a
/// workbook.
pub trait WriteContextLifecycle: WriteContext {
    /// Persists or discards pending output and releases owned resources.
    ///
    /// `on_exception` follows Java semantics: pending workbook bytes are
    /// discarded unless `writeExcelOnException` was enabled.
    ///
    /// # Errors
    ///
    /// Returns an output, handler, finalization, or stream-close error.
    fn finish_context(&mut self, on_exception: bool) -> Result<(), ExcelError>;
}

/// Holder surface exposed through [`WriteContext`].
pub trait WriteContextHolder {
    /// Returns the output path. (Java `WriteWorkbookHolder.getFile()`)
    fn path(&self) -> &Path;

    /// Returns the workbook-level handler context when available.
    /// (Java `WriteWorkbookHolder` via `WriteContextImpl.writeWorkbookHolder`)
    fn workbook_context(&self) -> Option<&WriteWorkbookContext> {
        None
    }

    /// Returns the active sheet handler context when available.
    /// (Java `WriteSheetHolder` via `WriteContextImpl.writeSheetHolder`)
    fn sheet_context(&self) -> Option<&WriteSheetContext> {
        None
    }

    /// Returns the zero-based table index when writing table content.
    /// (Java `WriteTableHolder.getTableNo()`)
    fn table_no(&self) -> Option<i32> {
        None
    }

    /// Returns the active sheet name when a sheet holder exists.
    fn sheet_name(&self) -> Option<&str> {
        self.sheet_context().map(WriteSheetContext::sheet_name)
    }

    /// Returns the resolved zero-based sheet number when known.
    fn sheet_no(&self) -> Option<i32> {
        self.sheet_context()
            .and_then(|context| context.write_sheet_holder().sheet_no())
    }

    /// Returns the latest physical row visible to the holder.
    fn last_row_index(&self) -> Option<u32> {
        self.sheet_context()
            .and_then(|context| context.write_sheet_holder().last_row_index())
    }

    /// Returns whether the active sheet has visible row data.
    fn has_data(&self) -> bool {
        self.sheet_context()
            .is_some_and(|context| context.write_sheet_holder().has_data())
    }

    /// Returns the active holder level. (Java `HolderEnum`)
    fn holder_type(&self) -> Holder;

    /// Returns the fully resolved header property.
    /// (Java `WriteHolder.excelWriteHeadProperty()`)
    fn excel_write_head_property(&self) -> &ExcelWriteHeadProperty;

    /// Returns the effective converter map for the active holder.
    /// (Java `ConfigurationHolder.converterMap()`)
    fn converter_map(&self) -> &ConverterRegistry;

    /// Returns whether this holder writes a header. (Java `needHead()`)
    fn need_head(&self) -> bool;

    /// Returns whether automatic header merging is enabled.
    fn automatic_merge_head(&self) -> bool;

    /// Returns the relative header row offset.
    fn relative_head_row_index(&self) -> i32;

    /// Returns whether include-list order controls output order.
    fn order_by_include_column(&self) -> bool;

    /// Returns included physical column indexes.
    fn include_column_indexes(&self) -> Option<&[usize]>;

    /// Returns included field names.
    fn include_column_field_names(&self) -> Option<&[String]>;

    /// Returns excluded physical column indexes.
    fn exclude_column_indexes(&self) -> &[usize];

    /// Returns excluded field names.
    fn exclude_column_field_names(&self) -> &[String];
}

/// Fully resolved Java `WriteHolder` state independent of a concrete backend.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteContextHolderState {
    /// Active holder level.
    pub holder_type: Holder,
    /// Resolved header metadata.
    pub excel_write_head_property: ExcelWriteHeadProperty,
    /// Effective workbook/sheet/table converter map.
    pub converter_map: ConverterRegistry,
    /// Whether a header is written.
    pub need_head: bool,
    /// Whether automatic header merging is enabled.
    pub automatic_merge_head: bool,
    /// Relative header row offset.
    pub relative_head_row_index: i32,
    /// Whether include-list order controls output.
    pub order_by_include_column: bool,
    /// Included physical columns.
    pub include_column_indexes: Option<Vec<usize>>,
    /// Included field names.
    pub include_column_field_names: Option<Vec<String>>,
    /// Excluded physical columns.
    pub exclude_column_indexes: Vec<usize>,
    /// Excluded field names.
    pub exclude_column_field_names: Vec<String>,
}

impl Default for WriteContextHolderState {
    fn default() -> Self {
        Self {
            holder_type: Holder::Workbook,
            excel_write_head_property: ExcelWriteHeadProperty::new(),
            converter_map: ConverterRegistry::default(),
            need_head: true,
            automatic_merge_head: true,
            relative_head_row_index: 0,
            order_by_include_column: false,
            include_column_indexes: None,
            include_column_field_names: None,
            exclude_column_indexes: Vec::new(),
            exclude_column_field_names: Vec::new(),
        }
    }
}

impl WriteContextHolderState {
    /// Clones the backend-neutral state exposed by a live Java-style holder.
    #[must_use]
    pub fn from_holder(holder: &dyn WriteContextHolder) -> Self {
        Self {
            holder_type: holder.holder_type(),
            excel_write_head_property: holder.excel_write_head_property().clone(),
            converter_map: holder.converter_map().clone(),
            need_head: holder.need_head(),
            automatic_merge_head: holder.automatic_merge_head(),
            relative_head_row_index: holder.relative_head_row_index(),
            order_by_include_column: holder.order_by_include_column(),
            include_column_indexes: holder.include_column_indexes().map(<[usize]>::to_vec),
            include_column_field_names: holder.include_column_field_names().map(<[String]>::to_vec),
            exclude_column_indexes: holder.exclude_column_indexes().to_vec(),
            exclude_column_field_names: holder.exclude_column_field_names().to_vec(),
        }
    }
}

/// Mirrors Java `WriteContextImpl implements WriteContext`.
///
/// Java owns POI workbook state; Rust exposes path and holder mirrors for
/// writer facades that delegate to `rust_xlsxwriter`.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteContextImpl {
    /// Output path. (Java `WriteWorkbookHolder.file`)
    path: PathBuf,
    /// Workbook-level handler context. (Java `WriteWorkbookHolder`)
    workbook_context: WriteWorkbookContext,
    /// Active sheet handler context. (Java `WriteSheetHolder`)
    sheet_context: Option<WriteSheetContext>,
    /// Active table index when writing table content. (Java `WriteTableHolder.tableNo`)
    table_no: Option<i32>,
    /// Resolved state of the current workbook/sheet/table holder.
    current_holder_state: WriteContextHolderState,
}

impl WriteContextImpl {
    /// Creates a write context bound to an output path.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        let path = path.into();
        Self {
            workbook_context: WriteWorkbookContext::new(&path),
            path,
            sheet_context: None,
            table_no: None,
            current_holder_state: WriteContextHolderState::default(),
        }
    }

    /// Returns the output path. (Java `WriteWorkbookHolder.getFile()`)
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Returns the workbook-level handler context.
    #[must_use]
    pub fn workbook_context(&self) -> &WriteWorkbookContext {
        &self.workbook_context
    }

    /// Returns the active sheet handler context, if any.
    #[must_use]
    pub fn sheet_context(&self) -> Option<&WriteSheetContext> {
        self.sheet_context.as_ref()
    }

    /// Returns the active table index, if any.
    #[must_use]
    pub const fn table_no(&self) -> Option<i32> {
        self.table_no
    }

    /// Returns the resolved current holder state.
    #[must_use]
    pub const fn current_holder_state(&self) -> &WriteContextHolderState {
        &self.current_holder_state
    }

    /// Replaces the resolved current holder state.
    pub fn set_current_holder_state(&mut self, state: WriteContextHolderState) {
        self.current_holder_state = state;
    }

    /// Updates the active sheet context. (Java `WriteContextImpl` sheet switch)
    pub fn set_sheet_context(&mut self, sheet_name: impl Into<String>) {
        self.sheet_context = Some(WriteSheetContext::new(sheet_name));
        self.current_holder_state.holder_type = Holder::Sheet;
    }

    /// Updates the active table index. (Java `WriteContextImpl` table switch)
    pub const fn set_table_no(&mut self, table_no: Option<i32>) {
        self.table_no = table_no;
        self.current_holder_state.holder_type = if table_no.is_some() {
            Holder::Table
        } else if self.sheet_context.is_some() {
            Holder::Sheet
        } else {
            Holder::Workbook
        };
    }
}

impl WriteContext for WriteContextImpl {
    fn current_write_holder(&self) -> &dyn WriteContextHolder {
        self
    }
}

impl WriteContextHolder for WriteContextImpl {
    fn path(&self) -> &Path {
        &self.path
    }

    fn workbook_context(&self) -> Option<&WriteWorkbookContext> {
        Some(&self.workbook_context)
    }

    fn sheet_context(&self) -> Option<&WriteSheetContext> {
        self.sheet_context.as_ref()
    }

    fn table_no(&self) -> Option<i32> {
        self.table_no
    }

    fn holder_type(&self) -> Holder {
        self.current_holder_state.holder_type
    }

    fn excel_write_head_property(&self) -> &ExcelWriteHeadProperty {
        &self.current_holder_state.excel_write_head_property
    }

    fn converter_map(&self) -> &ConverterRegistry {
        &self.current_holder_state.converter_map
    }

    fn need_head(&self) -> bool {
        self.current_holder_state.need_head
    }

    fn automatic_merge_head(&self) -> bool {
        self.current_holder_state.automatic_merge_head
    }

    fn relative_head_row_index(&self) -> i32 {
        self.current_holder_state.relative_head_row_index
    }

    fn order_by_include_column(&self) -> bool {
        self.current_holder_state.order_by_include_column
    }

    fn include_column_indexes(&self) -> Option<&[usize]> {
        self.current_holder_state.include_column_indexes.as_deref()
    }

    fn include_column_field_names(&self) -> Option<&[String]> {
        self.current_holder_state
            .include_column_field_names
            .as_deref()
    }

    fn exclude_column_indexes(&self) -> &[usize] {
        &self.current_holder_state.exclude_column_indexes
    }

    fn exclude_column_field_names(&self) -> &[String] {
        &self.current_holder_state.exclude_column_field_names
    }
}

/// Executes Java `WriteContext.finish(boolean onException)` semantics.
///
/// This function performs real dynamic dispatch to a resource-owning context;
/// it is not available for metadata-only [`WriteContextImpl`] values.
///
/// # Errors
///
/// Returns the concrete writer's output, handler, finalization, or close error.
pub fn finish_write_context(
    context: &mut dyn WriteContextLifecycle,
    on_exception: bool,
) -> Result<(), ExcelError> {
    context.finish_context(on_exception)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn write_context_impl_exposes_workbook_sheet_and_table_accessors() {
        let mut context = WriteContextImpl::new("out.xlsx");
        context.set_sheet_context("Sheet1");
        context.set_table_no(Some(2));

        let holder = context.current_write_holder();
        assert_eq!(holder.path(), Path::new("out.xlsx"));
        assert_eq!(
            holder.workbook_context().map(WriteWorkbookContext::path),
            Some(Path::new("out.xlsx"))
        );
        assert_eq!(
            holder.sheet_context().map(WriteSheetContext::sheet_name),
            Some("Sheet1")
        );
        assert_eq!(holder.table_no(), Some(2));
        assert_eq!(holder.holder_type(), Holder::Table);
    }

    #[derive(Default)]
    struct LifecycleProbe {
        on_exception: Option<bool>,
    }

    impl WriteContext for LifecycleProbe {
        fn current_write_holder(&self) -> &dyn WriteContextHolder {
            self
        }
    }

    impl WriteContextHolder for LifecycleProbe {
        fn path(&self) -> &Path {
            Path::new("probe.xlsx")
        }

        fn holder_type(&self) -> Holder {
            Holder::Workbook
        }

        fn excel_write_head_property(&self) -> &ExcelWriteHeadProperty {
            static PROPERTY: std::sync::OnceLock<ExcelWriteHeadProperty> =
                std::sync::OnceLock::new();
            PROPERTY.get_or_init(ExcelWriteHeadProperty::new)
        }

        fn converter_map(&self) -> &ConverterRegistry {
            static REGISTRY: std::sync::OnceLock<ConverterRegistry> = std::sync::OnceLock::new();
            REGISTRY.get_or_init(ConverterRegistry::default)
        }

        fn need_head(&self) -> bool {
            true
        }

        fn automatic_merge_head(&self) -> bool {
            true
        }

        fn relative_head_row_index(&self) -> i32 {
            0
        }

        fn order_by_include_column(&self) -> bool {
            false
        }

        fn include_column_indexes(&self) -> Option<&[usize]> {
            None
        }

        fn include_column_field_names(&self) -> Option<&[String]> {
            None
        }

        fn exclude_column_indexes(&self) -> &[usize] {
            &[]
        }

        fn exclude_column_field_names(&self) -> &[String] {
            &[]
        }
    }

    impl WriteContextLifecycle for LifecycleProbe {
        fn finish_context(&mut self, on_exception: bool) -> Result<(), ExcelError> {
            self.on_exception = Some(on_exception);
            Ok(())
        }
    }

    #[test]
    fn finish_write_context_dispatches_to_resource_owner() {
        let mut context = LifecycleProbe::default();
        finish_write_context(&mut context, true).expect("lifecycle should run");
        assert_eq!(context.on_exception, Some(true));
    }
}
