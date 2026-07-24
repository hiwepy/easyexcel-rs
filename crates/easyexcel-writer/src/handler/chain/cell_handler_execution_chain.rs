//! Mirrors Java `com.alibaba.excel.write.handler.chain.CellHandlerExecutionChain`.

use easyexcel_core::WriteCellContext;

/// Mirrors Java `CellHandlerExecutionChain` (a single linked-list node).
pub struct CellHandlerExecutionChain {
    pub(crate) handler: Option<Box<dyn easyexcel_core::WriteHandler>>,
    pub(crate) next: Option<Box<CellHandlerExecutionChain>>,
}

impl CellHandlerExecutionChain {
    /// Creates the head of an empty chain.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            handler: None,
            next: None,
        }
    }

    /// Creates a chain whose head contains `handler`. (Java constructor)
    #[must_use]
    pub fn with_handler(handler: Box<dyn easyexcel_core::WriteHandler>) -> Self {
        Self {
            handler: Some(handler),
            next: None,
        }
    }

    /// Appends a handler. (Java `addLast(WriteHandler)`)
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

    /// Runs the chain's cell lifecycle. (Java `beforeCellCreate`)
    pub fn before_cell_create(
        &mut self,
        context: &mut WriteCellContext,
    ) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.before_cell_create(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.before_cell_create(context)?;
        }
        Ok(())
    }

    /// Runs Java `afterCellCreate` in chain order.
    pub fn after_cell_create(&mut self, context: &WriteCellContext) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_cell_create(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_cell_create(context)?;
        }
        Ok(())
    }

    /// Runs Java `afterCellDataConverted` in chain order.
    pub fn after_cell_data_converted(
        &mut self,
        context: &WriteCellContext,
    ) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_cell_data_converted(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_cell_data_converted(context)?;
        }
        Ok(())
    }

    /// Runs Java `afterCellDispose` in chain order.
    pub fn after_cell_dispose(&mut self, context: &WriteCellContext) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_cell_dispose(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_cell_dispose(context)?;
        }
        Ok(())
    }
}

impl Default for CellHandlerExecutionChain {
    fn default() -> Self {
        Self::new()
    }
}
