//! Mirrors Java `com.alibaba.excel.write.handler.chain.RowHandlerExecutionChain`.

use easyexcel_core::WriteRowContext;

/// Mirrors Java `RowHandlerExecutionChain`.
pub struct RowHandlerExecutionChain {
    pub(crate) handler: Option<Box<dyn easyexcel_core::WriteHandler>>,
    pub(crate) next: Option<Box<RowHandlerExecutionChain>>,
}

impl RowHandlerExecutionChain {
    /// Creates an empty chain head.
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

    /// Runs Java `beforeRowCreate` in chain order.
    pub fn before_row_create(&mut self, context: &WriteRowContext) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.before_row_create(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.before_row_create(context)?;
        }
        Ok(())
    }

    /// Runs Java `afterRowCreate` in chain order.
    pub fn after_row_create(&mut self, context: &WriteRowContext) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_row_create(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_row_create(context)?;
        }
        Ok(())
    }

    /// Runs Java `afterRowDispose` in chain order.
    pub fn after_row_dispose(&mut self, context: &WriteRowContext) -> easyexcel_core::Result<()> {
        if let Some(handler) = self.handler.as_mut() {
            handler.after_row_dispose(context)?;
        }
        if let Some(next) = self.next.as_mut() {
            next.after_row_dispose(context)?;
        }
        Ok(())
    }
}

impl Default for RowHandlerExecutionChain {
    fn default() -> Self {
        Self::new()
    }
}
