//! Mirrors Java `com.alibaba.excel.write.merge.AbstractMergeStrategy`.

use easyexcel_core::{CellExtra, WriteCellContext, WriteHandler};

/// Mirrors Java `AbstractMergeStrategy implements CellWriteHandler`.
///
/// The Java side overrides `afterCellDispose` and calls the abstract
/// `merge(Sheet, Cell, Head, Integer relativeRowIndex)`. Rust mirrors the
/// structure so the strategy classes can override the hook.
pub trait AbstractMergeStrategy: WriteHandler {
    /// Called once per non-head cell. (Java `afterCellDispose`)
    fn after_cell_dispose(&mut self, context: &WriteCellContext) {
        let _ = context; // no-op default
    }

    /// Applies the merge to the worksheet. (Java `merge(Sheet, Cell, Head, Integer)`)
    fn merge(
        &mut self,
        sheet_name: &str,
        cell: &WriteCellContext,
        _extra: Option<&CellExtra>,
        _relative_row_index: Option<i32>,
    );
}
