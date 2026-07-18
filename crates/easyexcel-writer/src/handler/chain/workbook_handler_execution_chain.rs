//! Mirrors Java `com.alibaba.excel.write.handler.chain.WorkbookHandlerExecutionChain`.

use easyexcel_core::WriteWorkbookContext;

/// Mirrors Java `WorkbookHandlerExecutionChain`.
pub struct WorkbookHandlerExecutionChain {
    pub(crate) handler: Option<Box<dyn easyexcel_core::WriteHandler>>,
    pub(crate) next: Option<Box<WorkbookHandlerExecutionChain>>,
}

impl WorkbookHandlerExecutionChain {
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

    /// Runs the workbook lifecycle. (Java `afterWorkbookDispose`)
    pub fn after_workbook_dispose(
        &mut self,
        context: &WriteWorkbookContext,
    ) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_workbook(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_workbook_dispose(context)?;
        }
        Ok(())
    }
}

impl Default for WorkbookHandlerExecutionChain {
    fn default() -> Self {
        Self::new()
    }
}