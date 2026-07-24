//! Mirrors Java `com.alibaba.excel.write.handler.WriteHandler` plus its four
//! capability interfaces (Workbook/Sheet/Row/CellWriteHandler) collapsed into a
//! single trait.

use crate::analysis_context::Result;
use crate::cell_value::CellValue;
use crate::event::NotRepeatExecutor;
use crate::excel_cell_style::ExcelCellStyle;
use crate::metadata::property::LoopMergeProperty;
use crate::metadata::property::OnceAbsoluteMergeProperty;
use crate::write_cell_context::WriteCellContext;
use crate::write_row_context::WriteRowContext;
use crate::write_sheet_context::WriteSheetContext;
use crate::write_workbook_context::WriteWorkbookContext;

/// Intercepts the workbook, worksheet, row, and cell write lifecycle.
///
/// Java exposes four capability interfaces (`WorkbookWriteHandler`,
/// `SheetWriteHandler`, `RowWriteHandler`, `CellWriteHandler`) plus four
/// `Abstract*WriteHandler` skeletons. Rust collapses the eight lifecycle
/// hooks into a single trait; default implementations are no-ops so a
/// minimal handler only overrides the events it cares about.
///
/// Style / dimension strategies (Java `AbstractCellStyleStrategy`,
/// `AbstractColumnWidthStyleStrategy`, `AbstractRowHeightStyleStrategy`,
/// `LongestMatchColumnWidthStyleStrategy`) also override the optional
/// `style_*` accessors so `Box<dyn WriteHandler>` can apply them without
/// downcasting.
///
/// Any callback may return an error to stop the write immediately.
#[allow(clippy::missing_errors_doc)]
pub trait WriteHandler {
    /// Lower orders execute first. (Java `Order.order()`)
    fn order(&self) -> i32 {
        0
    }

    /// Exposes Java's optional [`NotRepeatExecutor`] capability.
    ///
    /// Rust trait objects cannot discover whether a handler implements a
    /// second trait at runtime. A handler that also implements
    /// [`NotRepeatExecutor`] therefore returns `Some(self)` here. The writer
    /// then keeps only the first handler with the same `unique_value()` after
    /// sorting by [`Self::order`], matching Java `AbstractWriteHolder`.
    fn as_not_repeat_executor(&self) -> Option<&dyn NotRepeatExecutor> {
        None
    }

    /// Called before the workbook is created. (Java `WorkbookWriteHandler.beforeWorkbookCreate`)
    fn before_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `WorkbookWriteHandler.beforeWorkbookCreate` lifecycle hook.
    ///
    /// The default delegates to the original Rust compatibility hook
    /// [`Self::before_workbook`].
    fn before_workbook_create(&mut self, context: &WriteWorkbookContext) -> Result<()> {
        self.before_workbook(context)
    }

    /// Exact Java `WorkbookWriteHandler.afterWorkbookCreate` lifecycle hook.
    fn after_workbook_create(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Called after the workbook is saved. (Java `WorkbookWriteHandler.afterWorkbookDispose`)
    fn after_workbook(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `WorkbookWriteHandler.afterWorkbookDispose` lifecycle hook.
    ///
    /// The default delegates to the original Rust compatibility hook
    /// [`Self::after_workbook`].
    fn after_workbook_dispose(&mut self, context: &WriteWorkbookContext) -> Result<()> {
        self.after_workbook(context)
    }

    /// Called after the worksheet is configured and before rows are written.
    /// (Java `SheetWriteHandler.beforeSheetCreate`)
    fn before_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `SheetWriteHandler.beforeSheetCreate` lifecycle hook.
    fn before_sheet_create(&mut self, context: &WriteSheetContext) -> Result<()> {
        self.before_sheet(context)
    }

    /// Exact Java `SheetWriteHandler.afterSheetCreate` lifecycle hook.
    fn after_sheet_create(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Called after all worksheet rows are written. (Java `SheetWriteHandler.afterSheetCreate`)
    fn after_sheet(&mut self, _context: &WriteSheetContext) -> Result<()> {
        Ok(())
    }

    /// Rust compatibility hook fired after all rows for a newly created sheet.
    ///
    /// Java has no separate sheet-dispose callback; its
    /// `afterSheetCreate` runs before row writing. The default preserves the
    /// historical Rust [`Self::after_sheet`] timing.
    fn after_sheet_dispose(&mut self, context: &WriteSheetContext) -> Result<()> {
        self.after_sheet(context)
    }

    /// Called before a row is written. (Java `RowWriteHandler.beforeRowCreate`)
    fn before_row(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `RowWriteHandler.beforeRowCreate` lifecycle hook.
    fn before_row_create(&mut self, context: &WriteRowContext) -> Result<()> {
        self.before_row(context)
    }

    /// Exact Java `RowWriteHandler.afterRowCreate` lifecycle hook.
    fn after_row_create(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Called after a row is written. (Java `RowWriteHandler.afterRowCreate`)
    fn after_row(&mut self, _context: &WriteRowContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `RowWriteHandler.afterRowDispose` lifecycle hook.
    fn after_row_dispose(&mut self, context: &WriteRowContext) -> Result<()> {
        self.after_row(context)
    }

    /// Called before a cell is written and may transform or skip it.
    /// (Java `CellWriteHandler.beforeCellCreate`)
    fn before_cell(&mut self, _context: &mut WriteCellContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `CellWriteHandler.beforeCellCreate` lifecycle hook.
    fn before_cell_create(&mut self, context: &mut WriteCellContext) -> Result<()> {
        self.before_cell(context)
    }

    /// Exact Java `CellWriteHandler.afterCellCreate` lifecycle hook.
    fn after_cell_create(&mut self, _context: &WriteCellContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `CellWriteHandler.afterCellDataConverted` lifecycle hook.
    fn after_cell_data_converted(&mut self, _context: &WriteCellContext) -> Result<()> {
        Ok(())
    }

    /// Called after a cell has been processed. (Java `CellWriteHandler.afterCellDispose`)
    fn after_cell(&mut self, _context: &WriteCellContext) -> Result<()> {
        Ok(())
    }

    /// Exact Java `CellWriteHandler.afterCellDispose` lifecycle hook.
    fn after_cell_dispose(&mut self, context: &WriteCellContext) -> Result<()> {
        self.after_cell(context)
    }

    /// Optional cell style from a registered style strategy.
    ///
    /// Mirrors Java `AbstractCellStyleStrategy` / `HorizontalCellStyleStrategy`
    /// / `AbstractVerticalCellStyleStrategy` applying `WriteCellStyle` during
    /// `afterCellDispose`. The XLSX write path merges non-`None` results into
    /// the cell format.
    fn style_cell_style(&self, _context: &WriteCellContext) -> Option<ExcelCellStyle> {
        None
    }

    /// Optional fixed column width from a registered width strategy.
    ///
    /// Mirrors Java `AbstractColumnWidthStyleStrategy.columnWidth` /
    /// `SimpleColumnWidthStyleStrategy`.
    fn style_column_width(&self, _column_index: usize) -> Option<u16> {
        None
    }

    /// Optional head row height from a registered row-height strategy.
    ///
    /// Mirrors Java `AbstractRowHeightStyleStrategy.setHeadColumnHeight` /
    /// `SimpleRowHeightStyleStrategy`.
    fn style_head_row_height(&self) -> Option<u16> {
        None
    }

    /// Optional content row height from a registered row-height strategy.
    ///
    /// Mirrors Java `AbstractRowHeightStyleStrategy.setContentColumnHeight` /
    /// `SimpleRowHeightStyleStrategy`.
    fn style_content_row_height(&self) -> Option<u16> {
        None
    }

    /// Whether the handler requests worksheet autofit as a column-width fallback.
    ///
    /// Mirrors an optional Rust extension around Java
    /// `LongestMatchColumnWidthStyleStrategy` (default path uses content
    /// byte-length + `setColumnWidth`; autofit remains opt-in).
    fn style_auto_column_width(&self) -> bool {
        false
    }

    /// Optional absolute merge region from a registered merge strategy.
    ///
    /// Mirrors Java `OnceAbsoluteMergeStrategy.afterSheetCreate` when the
    /// strategy is registered via `register_write_handler` (annotation
    /// `@OnceAbsoluteMerge` remains a separate metadata path).
    fn style_once_absolute_merge(&self) -> Option<OnceAbsoluteMergeProperty> {
        None
    }

    /// Optional repeating merge generated by `@ContentLoopMerge` or an
    /// explicitly registered Java-compatible `LoopMergeStrategy`.
    ///
    /// The tuple contains the merge dimensions and the resolved physical
    /// column index. Backends consume this after row disposal, matching Java's
    /// `LoopMergeStrategy.afterRowDispose`.
    fn style_loop_merge(&self) -> Option<(LoopMergeProperty, usize)> {
        None
    }
}

// `CellValue` import retained for downstream conversions; suppress unused warning.
#[allow(dead_code)]
fn _import_marker(v: CellValue) {
    let _ = v;
}
