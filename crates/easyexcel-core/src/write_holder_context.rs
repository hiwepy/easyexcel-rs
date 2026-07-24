//! Backend-neutral read-only views of Java write holders.

use std::path::{Path, PathBuf};

use crate::{
    ExcelWriteHeadProperty, Holder, WriteContext, WriteContextHolder, WriteContextHolderState,
};

/// Read-only runtime view of Java `WriteWorkbookHolder`.
///
/// The view deliberately exposes logical EasyExcel state rather than a fake
/// Apache POI workbook. Backend objects remain owned by the writer engine.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteWorkbookHolderView {
    path: PathBuf,
}

impl WriteWorkbookHolderView {
    /// Creates a workbook holder view for the active output.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        Self { path: path.into() }
    }

    /// Returns the active output path. (Java `WriteWorkbookHolder.getFile()`)
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }
}

/// Read-only runtime view of Java `WriteSheetHolder`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteSheetHolderView {
    sheet_name: String,
    sheet_no: Option<i32>,
    last_row_index: Option<u32>,
    has_data: bool,
}

impl WriteSheetHolderView {
    /// Creates a view for an active worksheet.
    #[must_use]
    pub fn new(sheet_name: impl Into<String>) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            sheet_no: None,
            last_row_index: None,
            has_data: false,
        }
    }

    /// Records the resolved zero-based sheet number.
    #[must_use]
    pub const fn with_sheet_no(mut self, sheet_no: i32) -> Self {
        self.sheet_no = Some(sheet_no);
        self
    }

    /// Records the latest physical row visible at this callback stage.
    #[must_use]
    pub const fn with_last_row_index(mut self, last_row_index: u32) -> Self {
        self.last_row_index = Some(last_row_index);
        self.has_data = true;
        self
    }

    /// Returns the resolved worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Returns the resolved zero-based sheet number, when known.
    #[must_use]
    pub const fn sheet_no(&self) -> Option<i32> {
        self.sheet_no
    }

    /// Returns the latest physical row visible at this callback stage.
    #[must_use]
    pub const fn last_row_index(&self) -> Option<u32> {
        self.last_row_index
    }

    /// Returns whether a physical row is visible at this callback stage.
    #[must_use]
    pub const fn has_data(&self) -> bool {
        self.has_data
    }
}

/// Read-only runtime view of Java `WriteTableHolder`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteTableHolderView {
    table_no: i32,
    parent_sheet_name: String,
}

impl WriteTableHolderView {
    /// Creates a view for the active table and its parent sheet.
    #[must_use]
    pub fn new(table_no: i32, parent_sheet_name: impl Into<String>) -> Self {
        Self {
            table_no,
            parent_sheet_name: parent_sheet_name.into(),
        }
    }

    /// Returns the zero-based table number. (Java `WriteTableHolder.getTableNo()`)
    #[must_use]
    pub const fn table_no(&self) -> i32 {
        self.table_no
    }

    /// Returns the parent worksheet name.
    #[must_use]
    pub fn parent_sheet_name(&self) -> &str {
        &self.parent_sheet_name
    }
}

/// Holder set captured for a concrete write-handler callback.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteHolderContext {
    workbook: Option<WriteWorkbookHolderView>,
    sheet: Option<WriteSheetHolderView>,
    table: Option<WriteTableHolderView>,
    current_holder_state: WriteContextHolderState,
}

impl WriteHolderContext {
    /// Creates an empty holder set for compatibility constructors.
    #[must_use]
    pub fn new() -> Self {
        Self {
            workbook: None,
            sheet: None,
            table: None,
            current_holder_state: WriteContextHolderState::default(),
        }
    }

    /// Attaches the active workbook holder view.
    #[must_use]
    pub fn with_workbook(mut self, workbook: WriteWorkbookHolderView) -> Self {
        self.workbook = Some(workbook);
        self
    }

    /// Attaches the active sheet holder view.
    #[must_use]
    pub fn with_sheet(mut self, sheet: WriteSheetHolderView) -> Self {
        self.sheet = Some(sheet);
        self
    }

    /// Attaches the active table holder view.
    #[must_use]
    pub fn with_table(mut self, table: WriteTableHolderView) -> Self {
        self.table = Some(table);
        self
    }

    /// Attaches the fully resolved Java `currentWriteHolder()` state.
    #[must_use]
    pub fn with_current_holder_state(mut self, state: WriteContextHolderState) -> Self {
        self.current_holder_state = state;
        self
    }

    /// Captures all backend-neutral holder state from a live write context.
    #[must_use]
    pub fn from_write_context(context: &dyn WriteContext) -> Self {
        let holder = context.current_write_holder();
        let mut snapshot = Self::new()
            .with_workbook(WriteWorkbookHolderView::new(holder.path()))
            .with_current_holder_state(WriteContextHolderState::from_holder(holder));

        if let Some(sheet_name) = holder.sheet_name() {
            let mut sheet = WriteSheetHolderView::new(sheet_name);
            if let Some(sheet_no) = holder.sheet_no() {
                sheet = sheet.with_sheet_no(sheet_no);
            }
            if let Some(last_row_index) = holder.last_row_index() {
                sheet = sheet.with_last_row_index(last_row_index);
            }
            snapshot = snapshot.with_sheet(sheet);
            if let Some(table_no) = holder.table_no() {
                snapshot = snapshot.with_table(WriteTableHolderView::new(table_no, sheet_name));
            }
        }
        snapshot
    }

    /// Sets callback-specific sheet and optional latest-row state while
    /// preserving the live holder's resolved sheet number.
    #[must_use]
    pub fn with_callback_sheet(
        mut self,
        sheet_name: impl Into<String>,
        last_row_index: Option<u32>,
    ) -> Self {
        let sheet_name = sheet_name.into();
        let mut sheet = WriteSheetHolderView::new(&sheet_name);
        if let Some(sheet_no) = self.sheet.as_ref().and_then(WriteSheetHolderView::sheet_no) {
            sheet = sheet.with_sheet_no(sheet_no);
        }
        if let Some(last_row_index) = last_row_index {
            sheet = sheet.with_last_row_index(last_row_index);
        }
        self.sheet = Some(sheet);
        if let Some(table_no) = self.table.as_ref().map(WriteTableHolderView::table_no) {
            self.table = Some(WriteTableHolderView::new(table_no, sheet_name));
        }
        self
    }

    /// Returns the active workbook holder view.
    #[must_use]
    pub const fn workbook(&self) -> Option<&WriteWorkbookHolderView> {
        self.workbook.as_ref()
    }

    /// Returns the active sheet holder view.
    #[must_use]
    pub const fn sheet(&self) -> Option<&WriteSheetHolderView> {
        self.sheet.as_ref()
    }

    /// Returns the active table holder view.
    #[must_use]
    pub const fn table(&self) -> Option<&WriteTableHolderView> {
        self.table.as_ref()
    }

    /// Returns the active write holder through the Java-compatible context API.
    #[must_use]
    pub fn current_write_holder(&self) -> &dyn WriteContextHolder {
        self
    }
}

impl Default for WriteHolderContext {
    fn default() -> Self {
        Self::new()
    }
}

impl WriteContext for WriteHolderContext {
    fn current_write_holder(&self) -> &dyn WriteContextHolder {
        self
    }
}

impl WriteContextHolder for WriteHolderContext {
    fn path(&self) -> &Path {
        self.workbook
            .as_ref()
            .map_or_else(|| Path::new(""), WriteWorkbookHolderView::path)
    }

    fn table_no(&self) -> Option<i32> {
        self.table.as_ref().map(WriteTableHolderView::table_no)
    }

    fn sheet_name(&self) -> Option<&str> {
        self.sheet.as_ref().map(WriteSheetHolderView::sheet_name)
    }

    fn sheet_no(&self) -> Option<i32> {
        self.sheet.as_ref().and_then(WriteSheetHolderView::sheet_no)
    }

    fn last_row_index(&self) -> Option<u32> {
        self.sheet
            .as_ref()
            .and_then(WriteSheetHolderView::last_row_index)
    }

    fn has_data(&self) -> bool {
        self.sheet
            .as_ref()
            .is_some_and(WriteSheetHolderView::has_data)
    }

    fn holder_type(&self) -> Holder {
        self.current_holder_state.holder_type
    }

    fn excel_write_head_property(&self) -> &ExcelWriteHeadProperty {
        &self.current_holder_state.excel_write_head_property
    }

    fn converter_map(&self) -> &crate::ConverterRegistry {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellValue, WriteCellContext, WriteRowContext, WriteSheetContext};

    #[test]
    fn holder_views_preserve_real_workbook_sheet_row_and_table_state() {
        let workbook = WriteWorkbookHolderView::new("target.xlsx");
        let sheet =
            WriteSheetContext::new("Users").with_holder_context(workbook.clone(), 2, Some(7));
        assert_eq!(
            sheet
                .write_workbook_holder()
                .map(WriteWorkbookHolderView::path),
            Some(Path::new("target.xlsx"))
        );
        assert_eq!(sheet.write_sheet_holder().sheet_name(), "Users");
        assert_eq!(sheet.write_sheet_holder().sheet_no(), Some(2));
        assert_eq!(
            sheet
                .write_table_holder()
                .map(WriteTableHolderView::table_no),
            Some(7)
        );

        let row = WriteRowContext::new("Users", 42, Some(3), false).with_holder_context(
            workbook.clone(),
            2,
            Some(7),
        );
        assert_eq!(row.write_sheet_holder().last_row_index(), Some(42));
        assert!(row.write_sheet_holder().has_data());

        let cell = WriteCellContext::new("Users", 42, 1, CellValue::Int(9)).with_holder_context(
            workbook,
            2,
            Some(7),
        );
        assert_eq!(
            cell.write_workbook_holder()
                .map(WriteWorkbookHolderView::path),
            Some(Path::new("target.xlsx"))
        );
        assert_eq!(
            cell.write_table_holder()
                .map(WriteTableHolderView::parent_sheet_name),
            Some("Users")
        );
    }

    #[test]
    fn compatibility_contexts_report_unknown_state_as_absent() {
        let sheet = WriteSheetContext::new("Sheet1");
        assert!(sheet.write_workbook_holder().is_none());
        assert_eq!(sheet.write_sheet_holder().sheet_no(), None);
        assert!(sheet.write_table_holder().is_none());
    }
}
