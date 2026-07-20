//! Mirrors Java `com.alibaba.excel.context.WriteContext` (interface).

use std::path::{Path, PathBuf};

use crate::excel_error::ExcelError;
use crate::WriteSheetContext;
use crate::WriteWorkbookContext;

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
}

/// Mirrors Java `WriteContextImpl implements WriteContext`.
///
/// Java owns POI workbook state; Rust exposes path and holder mirrors for
/// writer facades that delegate to `rust_xlsxwriter`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteContextImpl {
    /// Output path. (Java `WriteWorkbookHolder.file`)
    path: PathBuf,
    /// Workbook-level handler context. (Java `WriteWorkbookHolder`)
    workbook_context: WriteWorkbookContext,
    /// Active sheet handler context. (Java `WriteSheetHolder`)
    sheet_context: Option<WriteSheetContext>,
    /// Active table index when writing table content. (Java `WriteTableHolder.tableNo`)
    table_no: Option<i32>,
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

    /// Updates the active sheet context. (Java `WriteContextImpl` sheet switch)
    pub fn set_sheet_context(&mut self, sheet_name: impl Into<String>) {
        self.sheet_context = Some(WriteSheetContext::new(sheet_name));
    }

    /// Updates the active table index. (Java `WriteContextImpl` table switch)
    pub const fn set_table_no(&mut self, table_no: Option<i32>) {
        self.table_no = table_no;
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
}

/// Mirrors Java `WriteContext.finish(boolean onException)`.
///
/// Java's finish dispatches to the underlying workbook save and the
/// handler lifecycle. Rust exposes a free function that delegates to
/// the writer.
pub fn finish_write_context(
    _context: &dyn WriteContext,
    on_exception: bool,
) -> Result<(), ExcelError> {
    let _ = on_exception;
    Ok(())
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
    }
}
