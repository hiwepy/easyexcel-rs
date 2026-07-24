//! Runtime implementation of Java `com.alibaba.excel.util.WriteHandlerUtils`.
//!
//! Java centralizes context construction and handler-chain dispatch in this
//! class. Rust uses backend-neutral contexts but preserves the same lifecycle
//! stages and error propagation.

use std::path::PathBuf;

use crate::{
    CellValue, ExcelColumn, ExcelError, Result, WriteCellContext, WriteContext, WriteHandler,
    WriteRowContext, WriteSheetContext, WriteWorkbookContext,
};

/// Creates a workbook callback context from the live Java-style write context.
#[must_use]
pub fn create_workbook_write_handler_context(
    write_context: &dyn WriteContext,
) -> WriteWorkbookContext {
    WriteWorkbookContext::from_write_context(write_context)
}

/// Creates a compatibility workbook callback context from an output path.
#[must_use]
pub fn create_workbook_write_handler_context_from_path(
    path: impl Into<PathBuf>,
) -> WriteWorkbookContext {
    WriteWorkbookContext::new(path)
}

/// Dispatches Java `beforeWorkbookCreate`.
pub fn before_workbook_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers {
        handler.before_workbook_create(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterWorkbookCreate`.
pub fn after_workbook_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_workbook_create(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterWorkbookDispose`.
pub fn after_workbook_dispose(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteWorkbookContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_workbook_dispose(context)?;
    }
    Ok(())
}

/// Creates a sheet callback context from the live Java-style write context.
///
/// # Errors
///
/// Returns an error when no sheet has been selected yet.
pub fn create_sheet_write_handler_context(
    write_context: &dyn WriteContext,
) -> Result<WriteSheetContext> {
    WriteSheetContext::from_write_context(write_context)
        .ok_or_else(|| ExcelError::Format("write context has no active sheet holder".to_owned()))
}

/// Creates a compatibility sheet callback context from a worksheet name.
#[must_use]
pub fn create_sheet_write_handler_context_from_name(
    sheet_name: impl Into<String>,
) -> WriteSheetContext {
    WriteSheetContext::new(sheet_name)
}

/// Dispatches Java `beforeSheetCreate`.
pub fn before_sheet_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteSheetContext,
) -> Result<()> {
    for handler in handlers {
        handler.before_sheet_create(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterSheetCreate`.
pub fn after_sheet_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteSheetContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_sheet_create(context)?;
    }
    Ok(())
}

/// Creates a row callback context from the live Java-style write context.
///
/// # Errors
///
/// Returns an error when no sheet has been selected yet.
pub fn create_row_write_handler_context(
    write_context: &dyn WriteContext,
    row_index: u32,
    relative_row_index: Option<usize>,
    is_head: bool,
) -> Result<WriteRowContext> {
    let sheet_name = write_context
        .current_write_holder()
        .sheet_name()
        .ok_or_else(|| ExcelError::Format("write context has no active sheet holder".to_owned()))?;
    Ok(
        WriteRowContext::new(sheet_name, row_index, relative_row_index, is_head)
            .with_write_context(write_context),
    )
}

/// Creates a compatibility row callback context from a worksheet name.
#[must_use]
pub fn create_row_write_handler_context_from_sheet(
    sheet_name: impl Into<String>,
    row_index: u32,
    relative_row_index: Option<usize>,
    is_head: bool,
) -> WriteRowContext {
    WriteRowContext::new(sheet_name, row_index, relative_row_index, is_head)
}

/// Dispatches Java `beforeRowCreate`.
pub fn before_row_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteRowContext,
) -> Result<()> {
    for handler in handlers {
        handler.before_row_create(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterRowCreate`.
pub fn after_row_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteRowContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_row_create(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterRowDispose`.
pub fn after_row_dispose(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteRowContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_row_dispose(context)?;
    }
    Ok(())
}

/// Creates a cell callback context from the live Java-style write context.
///
/// # Errors
///
/// Returns an error when no sheet has been selected yet.
#[allow(clippy::too_many_arguments)]
pub fn create_cell_write_handler_context(
    write_context: &dyn WriteContext,
    row_index: u32,
    column_index: u16,
    relative_row_index: Option<usize>,
    is_head: bool,
    head_name: Option<String>,
    column: Option<&'static ExcelColumn>,
    value: CellValue,
) -> Result<WriteCellContext> {
    let sheet_name = write_context
        .current_write_holder()
        .sheet_name()
        .ok_or_else(|| ExcelError::Format("write context has no active sheet holder".to_owned()))?;
    Ok(create_cell_write_handler_context_from_sheet(
        sheet_name,
        row_index,
        column_index,
        relative_row_index,
        is_head,
        head_name,
        column,
        value,
    )
    .with_write_context(write_context))
}

/// Creates a compatibility cell callback context from a worksheet name.
#[must_use]
#[allow(clippy::too_many_arguments)]
pub fn create_cell_write_handler_context_from_sheet(
    sheet_name: impl Into<String>,
    row_index: u32,
    column_index: u16,
    relative_row_index: Option<usize>,
    is_head: bool,
    head_name: Option<String>,
    column: Option<&'static ExcelColumn>,
    value: CellValue,
) -> WriteCellContext {
    let mut context = WriteCellContext::new(sheet_name, row_index, column_index, value)
        .with_relative_row_index(relative_row_index);
    if let Some(column) = column {
        context = context.with_column(column);
    }
    if is_head {
        context = context.with_head(head_name.unwrap_or_default());
    }
    context
}

/// Dispatches Java `beforeCellCreate`.
pub fn before_cell_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &mut WriteCellContext,
) -> Result<()> {
    for handler in handlers {
        handler.before_cell_create(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterCellCreate`.
pub fn after_cell_create(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteCellContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_cell_create(context)?;
    }
    Ok(())
}

/// Finalizes converted cell metadata and dispatches Java
/// `afterCellDataConverted` for content cells.
pub fn after_cell_data_converted(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &mut WriteCellContext,
) -> Result<()> {
    context.activate_original_value();
    context.refresh_converted_data();
    if context.is_head {
        return Ok(());
    }
    for handler in handlers {
        handler.after_cell_data_converted(context)?;
    }
    Ok(())
}

/// Dispatches Java `afterCellDispose`.
pub fn after_cell_dispose(
    handlers: &mut [Box<dyn WriteHandler>],
    context: &WriteCellContext,
) -> Result<()> {
    for handler in handlers {
        handler.after_cell_dispose(context)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    };

    use super::*;

    struct Probe(Arc<AtomicUsize>);

    impl WriteHandler for Probe {
        fn before_workbook_create(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.0.fetch_add(1, Ordering::Relaxed);
            Ok(())
        }
    }

    #[test]
    fn utilities_dispatch_real_handler_contexts() {
        let calls = Arc::new(AtomicUsize::new(0));
        let mut handlers: Vec<Box<dyn WriteHandler>> = vec![Box::new(Probe(Arc::clone(&calls)))];
        let context = create_workbook_write_handler_context_from_path("real.xlsx");
        before_workbook_create(&mut handlers, &context).expect("dispatch");
        assert_eq!(context.path().to_string_lossy(), "real.xlsx");
        assert_eq!(calls.load(Ordering::Relaxed), 1);
    }

    #[test]
    fn java_style_context_factories_clone_the_live_holder_graph() -> Result<()> {
        let live_context = crate::WriteHolderContext::new()
            .with_workbook(crate::WriteWorkbookHolderView::new("live.xlsx"))
            .with_sheet(
                crate::WriteSheetHolderView::new("Users")
                    .with_sheet_no(4)
                    .with_last_row_index(9),
            )
            .with_table(crate::WriteTableHolderView::new(3, "Users"))
            .with_current_holder_state(crate::WriteContextHolderState {
                holder_type: crate::Holder::Table,
                excel_write_head_property: crate::ExcelWriteHeadProperty::from_head(
                    None,
                    None,
                    Some(vec![vec!["Name".to_owned()]]),
                    crate::ExcelWriteMetadata::default(),
                ),
                converter_map: crate::ConverterRegistry::default(),
                need_head: false,
                automatic_merge_head: false,
                relative_head_row_index: 2,
                order_by_include_column: true,
                include_column_indexes: Some(vec![2, 0]),
                include_column_field_names: Some(vec!["name".to_owned()]),
                exclude_column_indexes: vec![7],
                exclude_column_field_names: vec!["secret".to_owned()],
            });

        let workbook = create_workbook_write_handler_context(&live_context);
        assert_eq!(workbook.path(), std::path::Path::new("live.xlsx"));
        assert_eq!(
            workbook
                .write_context()
                .current_write_holder()
                .holder_type(),
            crate::Holder::Table
        );

        let sheet = create_sheet_write_handler_context(&live_context)?;
        assert_eq!(sheet.sheet_name(), "Users");
        assert_eq!(sheet.write_sheet_holder().sheet_no(), Some(4));
        assert_eq!(
            sheet
                .write_table_holder()
                .map(crate::WriteTableHolderView::table_no),
            Some(3)
        );

        let row = create_row_write_handler_context(&live_context, 12, Some(5), false)?;
        assert_eq!(row.write_sheet_holder().last_row_index(), Some(12));
        assert_eq!(
            row.write_context()
                .current_write_holder()
                .include_column_indexes(),
            Some(&[2, 0][..])
        );

        let cell = create_cell_write_handler_context(
            &live_context,
            12,
            2,
            Some(5),
            false,
            None,
            None,
            CellValue::String("Ada".to_owned()),
        )?;
        let holder = cell.write_context().current_write_holder();
        assert_eq!(holder.table_no(), Some(3));
        assert!(!holder.need_head());
        assert!(!holder.automatic_merge_head());
        assert_eq!(holder.relative_head_row_index(), 2);
        assert_eq!(
            holder
                .excel_write_head_property()
                .head_map()
                .get(&0)
                .map(|head| head.head_name_list()),
            Some(&["Name".to_owned()][..])
        );
        Ok(())
    }

    #[test]
    fn java_style_sheet_row_and_cell_factories_reject_missing_sheet() {
        let context = crate::WriteContextImpl::new("workbook-only.xlsx");
        assert!(create_sheet_write_handler_context(&context).is_err());
        assert!(create_row_write_handler_context(&context, 0, None, false).is_err());
        assert!(
            create_cell_write_handler_context(
                &context,
                0,
                0,
                None,
                false,
                None,
                None,
                CellValue::Empty,
            )
            .is_err()
        );
    }
}
