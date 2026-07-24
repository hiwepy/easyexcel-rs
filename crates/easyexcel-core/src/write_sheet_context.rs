//! Mirrors Java `com.alibaba.excel.write.handler.context.SheetWriteHandlerContext`.

use crate::WriteContext;

/// Worksheet-level write lifecycle context.
///
/// Mirrors Java `SheetWriteHandlerContext` (`writeSheetHolder.getSheetName()`).
#[derive(Debug, Clone, PartialEq)]
pub struct WriteSheetContext {
    sheet_name: String,
    holders: WriteHolderContext,
}

impl WriteSheetContext {
    /// Returns this backend-neutral sheet context.
    #[must_use]
    pub const fn sheet(&self) -> &Self {
        self
    }

    /// Creates a worksheet context.
    #[must_use]
    pub fn new(sheet_name: impl Into<String>) -> Self {
        let sheet_name = sheet_name.into();
        Self {
            holders: WriteHolderContext::new().with_sheet(WriteSheetHolderView::new(&sheet_name)),
            sheet_name,
        }
    }

    /// Creates a sheet callback context from a live [`WriteContext`].
    ///
    /// Returns `None` before the context has selected a sheet.
    #[must_use]
    pub fn from_write_context(context: &dyn WriteContext) -> Option<Self> {
        let holder = context.current_write_holder();
        let sheet_name = holder.sheet_name()?.to_owned();
        Some(Self {
            holders: WriteHolderContext::from_write_context(context)
                .with_callback_sheet(&sheet_name, holder.last_row_index()),
            sheet_name,
        })
    }

    /// Replaces compatibility holder data with a live-context snapshot.
    #[must_use]
    pub fn with_write_context(mut self, context: &dyn WriteContext) -> Self {
        self.holders = WriteHolderContext::from_write_context(context)
            .with_callback_sheet(&self.sheet_name, None);
        self
    }

    /// Returns the worksheet name. (Java `WriteSheetHolder.getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Attaches the workbook, resolved sheet number, and optional table.
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

    /// Attaches all holder views and the resolved Java `currentWriteHolder()` state.
    #[must_use]
    pub fn with_resolved_holder_context(
        mut self,
        workbook: WriteWorkbookHolderView,
        sheet_no: i32,
        table_no: Option<i32>,
        current_holder_state: crate::WriteContextHolderState,
    ) -> Self {
        let sheet = WriteSheetHolderView::new(&self.sheet_name).with_sheet_no(sheet_no);
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

    /// Returns the active workbook holder view, when the writer supplied one.
    #[must_use]
    pub const fn write_workbook_holder(&self) -> Option<&WriteWorkbookHolderView> {
        self.holders.workbook()
    }

    /// Returns the active sheet holder view.
    #[must_use]
    pub fn write_sheet_holder(&self) -> &WriteSheetHolderView {
        self.holders
            .sheet()
            .expect("sheet contexts always carry a sheet holder")
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
    WriteHolderContext, WriteSheetHolderView, WriteTableHolderView, WriteWorkbookHolderView,
};
