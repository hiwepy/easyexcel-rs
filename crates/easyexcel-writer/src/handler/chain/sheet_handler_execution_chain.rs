//! Mirrors Java `com.alibaba.excel.write.handler.chain.SheetHandlerExecutionChain`.

use easyexcel_core::WriteSheetContext;

/// Mirrors Java `SheetHandlerExecutionChain`.
pub struct SheetHandlerExecutionChain {
    pub(crate) handler: Option<Box<dyn easyexcel_core::WriteHandler>>,
    pub(crate) next: Option<Box<SheetHandlerExecutionChain>>,
}

impl SheetHandlerExecutionChain {
    /// Creates an empty chain head.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            handler: None,
            next: None,
        }
    }

    /// Appends a handler. (Java `addLast`)
    pub fn add_last(&mut self, handler: Box<dyn easyexcel_core::WriteHandler>) {
        match self.next.as_mut() {
            Some(next) => next.add_last(handler),
            None => {
                self.next = Some(Box::new(Self {
                    handler: Some(handler),
                    next: None,
                }));
            }
        }
    }

    /// Runs the sheet lifecycle. (Java `afterSheetCreate`)
    pub fn after_sheet_create(&mut self, context: &WriteSheetContext) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_sheet(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_sheet_create(context)?;
        }
        Ok(())
    }
}

impl Default for SheetHandlerExecutionChain {
    fn default() -> Self {
        Self::new()
    }
}