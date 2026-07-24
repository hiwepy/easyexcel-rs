//! Mirrors Java `com.alibaba.excel.write.handler.context.WorkbookWriteHandlerContext`.

use std::path::{Path, PathBuf};

use crate::{WriteContext, WriteHolderContext, WriteWorkbookHolderView};

/// Workbook-level write lifecycle context.
///
/// Mirrors Java `WorkbookWriteHandlerContext` (`writeContext`,
/// `writeWorkbookHolder`). Rust collapses it to the logical path because the
/// `rust_xlsxwriter::Workbook` is held privately by the [`crate::ExcelWriter`].
#[derive(Debug, Clone, PartialEq)]
pub struct WriteWorkbookContext {
    path: PathBuf,
    holders: WriteHolderContext,
}

impl WriteWorkbookContext {
    /// Returns this backend-neutral workbook context.
    #[must_use]
    pub const fn workbook(&self) -> &Self {
        self
    }

    /// Creates a workbook context for an output path.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        let path = path.into();
        Self {
            holders: WriteHolderContext::new().with_workbook(WriteWorkbookHolderView::new(&path)),
            path,
        }
    }

    /// Creates the Java callback context from a live [`WriteContext`].
    #[must_use]
    pub fn from_write_context(context: &dyn WriteContext) -> Self {
        let holders = WriteHolderContext::from_write_context(context);
        let path = holders.current_write_holder().path().to_path_buf();
        Self { path, holders }
    }

    /// Returns the output path. (Java `WriteWorkbookHolder.getFile()`)
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }

    /// Returns the live workbook holder view carried by this callback.
    #[must_use]
    pub fn write_workbook_holder(&self) -> &WriteWorkbookHolderView {
        self.holders
            .workbook()
            .expect("workbook contexts always carry a workbook holder")
    }

    /// Returns all holder views captured for this callback.
    #[must_use]
    pub const fn write_context(&self) -> &WriteHolderContext {
        &self.holders
    }
}
