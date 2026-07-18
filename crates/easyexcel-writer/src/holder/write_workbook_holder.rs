//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder`.

use std::collections::HashMap;

use easyexcel_core::WriteHandler;
use crate::MirroredWriteSheetHolder as WriteSheetHolder;

/// Mirrors Java `WriteWorkbookHolder extends AbstractWriteHolder`.
///
/// The Java side aggregates the `rust_xlsxwriter::Workbook` POI handle, the
/// in-progress sheet holders, and the writer's handler list. Rust holds the
/// same data inside [`crate::ExcelWriter`]; this struct is provided for
/// parity so handler context builders can carry an `&WriteWorkbookHolder`
/// exactly as Java does.
pub struct WriteWorkbookHolder<'a> {
    path: String,
    sheets: HashMap<String, WriteSheetHolder<'a>>,
    handlers: Vec<Box<dyn WriteHandler>>,
}

impl<'a> WriteWorkbookHolder<'a> {
    /// Creates a holder matching the Java `WriteWorkbookHolder(WriteWorkbook)`
    /// initialiser.
    #[must_use]
    pub fn new(path: impl Into<String>) -> Self {
        Self {
            path: path.into(),
            sheets: HashMap::new(),
            handlers: Vec::new(),
        }
    }

    /// Returns the output path. (Java `getFile()`)
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the in-progress sheet holders. (Java `getHasBeenInitializedSheetNameMap()`)
    #[must_use]
    pub fn sheets(&self) -> &HashMap<String, WriteSheetHolder<'a>> {
        &self.sheets
    }

    /// Returns a mutable handle on the in-progress sheet holders.
    pub fn sheets_mut(&mut self) -> &mut HashMap<String, WriteSheetHolder<'a>> {
        &mut self.sheets
    }

    /// Returns the ordered write handler list. (Java `getWriteHandlerList()`)
    #[must_use]
    pub fn handlers(&self) -> &[Box<dyn WriteHandler>] {
        &self.handlers
    }

    /// Appends a handler. (Java `setWriteHandlerList` step)
    pub fn push_handler(&mut self, handler: Box<dyn WriteHandler>) {
        self.handlers.push(handler);
    }
}