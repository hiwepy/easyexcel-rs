//! Mirrors Java `com.alibaba.excel.write.handler.context.RowWriteHandlerContext`.

/// Row-level write lifecycle context.
///
/// Mirrors Java `RowWriteHandlerContext` (`writeSheetHolder`, `writeTableHolder`,
/// `rowIndex`, `relativeRowIndex`, `head`). Rust keeps only the fields a
/// handler needs and drops the `Row` POI object because `rust_xlsxwriter`
/// does not expose it for handler interception.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteRowContext {
    /// Worksheet name.
    pub sheet_name: String,
    /// Physical zero-based row index.
    pub row_index: u32,
    /// Relative index within the current head or content block.
    ///
    /// Mirrors Java `RowWriteHandlerContext.relativeRowIndex`.
    pub relative_row_index: Option<usize>,
    /// Whether this is a header row.
    pub is_head: bool,
    row: WriteRowHandle,
    holders: WriteHolderContext,
}

impl WriteRowContext {
    /// Returns the mutable logical row handle.
    #[must_use]
    pub const fn row(&self) -> &WriteRowHandle {
        &self.row
    }

    /// Creates a row handler context.
    #[must_use]
    pub fn new(
        sheet_name: impl Into<String>,
        row_index: u32,
        relative_row_index: Option<usize>,
        is_head: bool,
    ) -> Self {
        let sheet_name = sheet_name.into();
        Self {
            row: WriteRowHandle::new(row_index),
            holders: WriteHolderContext::new()
                .with_sheet(WriteSheetHolderView::new(&sheet_name).with_last_row_index(row_index)),
            sheet_name,
            row_index,
            relative_row_index,
            is_head,
        }
    }

    /// Attaches the real writer holder state visible for this row callback.
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
            .expect("row contexts always carry a sheet holder")
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
use crate::{
    WriteContext, WriteHolderContext, WriteRowHandle, WriteSheetHolderView, WriteTableHolderView,
    WriteWorkbookHolderView,
};
