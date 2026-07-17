//! Mirrors Java `com.alibaba.excel.write.handler.context.WorkbookWriteHandlerContext`.

use std::path::{Path, PathBuf};

/// Workbook-level write lifecycle context.
///
/// Mirrors Java `WorkbookWriteHandlerContext` (`writeContext`,
/// `writeWorkbookHolder`). Rust collapses it to the logical path because the
/// `rust_xlsxwriter::Workbook` is held privately by the [`crate::ExcelWriter`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteWorkbookContext {
    path: PathBuf,
}

impl WriteWorkbookContext {
    /// Creates a workbook context for an output path.
    #[must_use]
    pub fn new(path: impl Into<PathBuf>) -> Self {
        Self { path: path.into() }
    }

    /// Returns the output path. (Java `WriteWorkbookHolder.getFile()`)
    #[must_use]
    pub fn path(&self) -> &Path {
        &self.path
    }
}
