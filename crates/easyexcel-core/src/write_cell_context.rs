//! Mirrors Java `com.alibaba.excel.write.handler.context.CellWriteHandlerContext`.

use crate::cell_value::CellValue;
use crate::enum_cell_data_type::CellDataType;
use crate::excel_column::ExcelColumn;
use crate::{
    WriteCellHandle, WriteContext, WriteHolderContext, WriteSheetHolderView, WriteTableHolderView,
    WriteWorkbookHolderView,
};

/// Mutable cell-level write lifecycle context.
///
/// Mirrors Java `CellWriteHandlerContext` (13 fields). Rust keeps only the
/// fields a handler actually mutates and exposes `skip: bool` so handlers
/// can suppress writing a cell without juggling the underlying POI types.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellContext {
    /// Worksheet name.
    pub sheet_name: String,
    /// Physical zero-based row index.
    pub row_index: u32,
    /// Physical zero-based column index.
    pub column_index: u16,
    /// Rust field name, when backed by a typed column.
    pub field: Option<&'static str>,
    /// Resolved static head/content metadata for this typed field.
    pub column: Option<&'static ExcelColumn>,
    /// Header label at this level, when this is a header cell.
    pub head_name: Option<String>,
    /// Whether this is a header cell.
    pub is_head: bool,
    /// Relative row index within head or content (Java `relativeRowIndex`).
    ///
    /// Used by `HorizontalCellStyleStrategy` to cycle content styles.
    pub relative_row_index: Option<usize>,
    /// Value before write handlers transform it.
    ///
    /// Mirrors Java `CellWriteHandlerContext.originalValue`. Typed writer
    /// paths replace the constructor default with the field's value before a
    /// registered or annotation converter runs.
    pub original_value: Option<CellValue>,
    /// Declared Rust field type before conversion.
    ///
    /// Mirrors Java `CellWriteHandlerContext.originalFieldClass`.
    pub original_field_type: Option<&'static str>,
    /// Source value held until Java's conversion stage begins.
    pending_original_value: Option<CellValue>,
    /// Declared field type held until Java's conversion stage begins.
    pending_original_field_type: Option<&'static str>,
    /// Converted cell data visible from `afterCellDataConverted` onward.
    ///
    /// Java permits multiple `WriteCellData` values. The current typed writer
    /// emits one scalar value, but retains the list shape for handler parity.
    pub cell_data_list: Vec<CellValue>,
    /// Target cell type selected by conversion.
    pub target_cell_data_type: Option<CellDataType>,
    /// Suppresses annotation/strategy style filling for this cell.
    ///
    /// Mirrors Java `CellWriteHandlerContext.ignoreFillStyle`.
    pub ignore_fill_style: bool,
    /// Value that will be written. A handler may replace it.
    pub value: CellValue,
    /// A handler may set this to suppress the physical cell.
    pub skip: bool,
    cell: WriteCellHandle,
    holders: WriteHolderContext,
}

impl WriteCellContext {
    /// Returns the mutable logical cell handle.
    #[must_use]
    pub const fn cell(&self) -> &WriteCellHandle {
        &self.cell
    }

    /// Creates a cell handler context before cell conversion callbacks run.
    #[must_use]
    pub fn new(
        sheet_name: impl Into<String>,
        row_index: u32,
        column_index: u16,
        value: CellValue,
    ) -> Self {
        let sheet_name = sheet_name.into();
        Self {
            cell: WriteCellHandle::new(row_index, column_index, value.clone()),
            holders: WriteHolderContext::new()
                .with_sheet(WriteSheetHolderView::new(&sheet_name).with_last_row_index(row_index)),
            sheet_name,
            row_index,
            column_index,
            field: None,
            column: None,
            head_name: None,
            is_head: false,
            relative_row_index: None,
            original_value: None,
            original_field_type: None,
            pending_original_value: Some(value.clone()),
            pending_original_field_type: None,
            cell_data_list: Vec::new(),
            target_cell_data_type: None,
            ignore_fill_style: false,
            value,
            skip: false,
        }
    }

    /// Attaches typed column metadata.
    #[must_use]
    pub const fn with_column(mut self, column: &'static ExcelColumn) -> Self {
        self.field = if column.field.is_empty() {
            None
        } else {
            Some(column.field)
        };
        self.pending_original_field_type = column.field_type;
        self.column = Some(column);
        self
    }

    /// Replaces the source value captured before conversion.
    #[must_use]
    pub fn with_original_value(mut self, original_value: CellValue) -> Self {
        self.pending_original_value = Some(original_value);
        self
    }

    /// Clears the source value for header cells.
    ///
    /// Java does not assign `originalValue` while creating head rows.
    #[must_use]
    pub fn without_original_value(mut self) -> Self {
        self.original_value = None;
        self.original_field_type = None;
        self.pending_original_value = None;
        self.pending_original_field_type = None;
        self
    }

    /// Makes pre-converter metadata visible at Java's conversion stage.
    pub fn activate_original_value(&mut self) {
        self.original_value = self.pending_original_value.clone();
        self.original_field_type = self.pending_original_field_type;
    }

    /// Marks a header cell and records its current label.
    #[must_use]
    pub fn with_head(mut self, head_name: impl Into<String>) -> Self {
        self.is_head = true;
        self.head_name = Some(head_name.into());
        self
    }

    /// Sets the relative row index.
    #[must_use]
    pub const fn with_relative_row_index(mut self, relative_row_index: Option<usize>) -> Self {
        self.relative_row_index = relative_row_index;
        self
    }

    /// Returns the first converted cell value.
    ///
    /// Mirrors Java `CellWriteHandlerContext.getFirstCellData()`.
    #[must_use]
    pub fn first_cell_data(&self) -> Option<&CellValue> {
        self.cell_data_list.first()
    }

    /// Refreshes conversion metadata after a handler changes [`Self::value`].
    pub fn refresh_converted_data(&mut self) {
        self.target_cell_data_type = Some(self.value.data_type());
        self.cell_data_list.clear();
        self.cell_data_list.push(self.value.clone());
    }

    /// Applies mutations requested through [`Self::cell`].
    ///
    /// Writer backends call this after the logical callback chain and before
    /// committing the physical cell.
    pub fn apply_cell_mutations(&mut self) {
        if let Some(value) = self.cell.requested_value() {
            self.value = value;
            self.refresh_converted_data();
        }
        if let Some(skip) = self.cell.requested_skip() {
            self.skip = skip;
        }
    }

    /// Synchronizes the logical handle after compatibility callbacks mutate
    /// [`Self::value`] directly.
    pub fn sync_cell_handle(&self) {
        self.cell.sync_value(&self.value);
    }

    /// Attaches the real writer holder state visible for this cell callback.
    #[must_use]
    pub fn with_holder_context(
        mut self,
        workbook: WriteWorkbookHolderView,
        sheet_no: i32,
        table_no: Option<i32>,
    ) -> Self {
        let holder_type = if table_no.is_some() {
            crate::Holder::Table
        } else {
            crate::Holder::Sheet
        };
        self = self.with_resolved_holder_context(
            workbook,
            sheet_no,
            table_no,
            crate::WriteContextHolderState {
                holder_type,
                ..crate::WriteContextHolderState::default()
            },
        );
        self
    }

    /// Replaces compatibility holder data with a live-context snapshot.
    #[must_use]
    pub fn with_write_context(mut self, context: &dyn WriteContext) -> Self {
        self.holders = WriteHolderContext::from_write_context(context)
            .with_callback_sheet(&self.sheet_name, Some(self.row_index));
        self
    }

    /// Attaches all holder views and the resolved Java `currentWriteHolder()` state.
    #[must_use]
    pub fn with_resolved_holder_context(
        mut self,
        workbook: WriteWorkbookHolderView,
        sheet_no: i32,
        table_no: Option<i32>,
        current_holder_state: crate::WriteContextHolderState,
    ) -> Self {
        let sheet = WriteSheetHolderView::new(&self.sheet_name)
            .with_sheet_no(sheet_no)
            .with_last_row_index(self.row_index);
        self.holders = WriteHolderContext::new()
            .with_workbook(workbook)
            .with_sheet(sheet)
            .with_current_holder_state(current_holder_state);
        if let Some(table_no) = table_no {
            self.holders = self
                .holders
                .with_table(WriteTableHolderView::new(table_no, &self.sheet_name));
        }
        self
    }

    /// Returns the active workbook holder view, when supplied by the writer.
    #[must_use]
    pub const fn write_workbook_holder(&self) -> Option<&WriteWorkbookHolderView> {
        self.holders.workbook()
    }

    /// Returns the active sheet holder view.
    #[must_use]
    pub fn write_sheet_holder(&self) -> &WriteSheetHolderView {
        self.holders
            .sheet()
            .expect("cell contexts always carry a sheet holder")
    }

    /// Returns the active table holder view for table callbacks.
    #[must_use]
    pub const fn write_table_holder(&self) -> Option<&WriteTableHolderView> {
        self.holders.table()
    }

    /// Returns all holder views captured for this callback.
    #[must_use]
    pub const fn write_context(&self) -> &WriteHolderContext {
        &self.holders
    }
}
