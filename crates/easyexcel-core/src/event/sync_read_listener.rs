//! Mirrors Java `com.alibaba.excel.event.SyncReadListener`.

use crate::analysis_context::AnalysisContext;
use crate::read_listener::ReadListener;

/// Synchronous data reading.
///
/// Rust port of Java `SyncReadListener extends AnalysisEventListener<Object>`.
/// Java collects every row into a `List<Object>`. The Rust port mirrors
/// the same buffer so `doReadAllSync()` callers can retrieve the list.
pub struct SyncReadListener {
    list: Vec<crate::CellValue>,
}

impl SyncReadListener {
    /// Creates an empty listener.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            list: Vec::new(),
        }
    }

    /// Returns the collected list. (Java `getList()`)
    #[must_use]
    pub fn list(&self) -> &[crate::CellValue] {
        &self.list
    }

    /// Sets the list. (Java `setList(List)`)
    pub fn set_list(&mut self, list: Vec<crate::CellValue>) {
        self.list = list;
    }
}

impl Default for SyncReadListener {
    fn default() -> Self {
        Self::new()
    }
}

impl ReadListener<crate::CellValue> for SyncReadListener {
    fn invoke(
        &mut self,
        data: crate::CellValue,
        _context: &AnalysisContext,
    ) -> crate::analysis_context::Result<()> {
        self.list.push(data);
        Ok(())
    }
}
