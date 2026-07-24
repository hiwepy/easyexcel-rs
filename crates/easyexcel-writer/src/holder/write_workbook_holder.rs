//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder`.

use std::collections::HashMap;
use std::ops::{Deref, DerefMut};

use crate::MirroredWriteSheetHolder as WriteSheetHolder;
use crate::holder::abstract_write_holder::AbstractWriteHolder;
use crate::metadata::WriteBasicParameter;
use easyexcel_core::WriteHandler;

/// Mirrors Java `WriteWorkbookHolder extends AbstractWriteHolder`.
///
/// The Java side aggregates the `rust_xlsxwriter::Workbook` POI handle, the
/// in-progress sheet holders, and the writer's handler list. Rust holds the
/// same data inside [`crate::ExcelWriter`]; this owned builder-side mirror is
/// retained for Java package/API parity. Runtime callbacks expose the actual
/// logical state through [`easyexcel_core::WriteWorkbookHolderView`].
pub struct WriteWorkbookHolder<'a> {
    abstract_holder: AbstractWriteHolder,
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
            abstract_holder: AbstractWriteHolder::default(),
            path: path.into(),
            sheets: HashMap::new(),
            handlers: Vec::new(),
        }
    }

    /// Creates a workbook holder from nullable write parameters.
    #[must_use]
    pub fn from_parameter(path: impl Into<String>, parameter: &WriteBasicParameter) -> Self {
        let mut holder = Self::new(path);
        holder.abstract_holder = AbstractWriteHolder::from_parameter(parameter, None);
        holder
    }

    /// Returns the inherited write-holder state.
    #[must_use]
    pub const fn abstract_holder(&self) -> &AbstractWriteHolder {
        &self.abstract_holder
    }

    /// Returns mutable inherited write-holder state.
    pub const fn abstract_holder_mut(&mut self) -> &mut AbstractWriteHolder {
        &mut self.abstract_holder
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

impl Deref for WriteWorkbookHolder<'_> {
    type Target = AbstractWriteHolder;

    fn deref(&self) -> &Self::Target {
        &self.abstract_holder
    }
}

impl DerefMut for WriteWorkbookHolder<'_> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.abstract_holder
    }
}
