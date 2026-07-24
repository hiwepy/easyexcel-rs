//! Mirrors Java `com.alibaba.excel.read.listener.ReadListener<T>` (and the
//! `IgnoreExceptionReadListener` default implementation).

use std::collections::HashMap;

use crate::analysis_context::{AnalysisContext, ErrorAction, Result};
use crate::cell_extra::CellExtra;
use crate::excel_error::ExcelError;

/// Event listener equivalent to Java `EasyExcel`'s `ReadListener`.
///
/// Java `ReadListener` is an interface with one abstract method (`invoke`).
/// Rust keeps the same shape: `invoke` is the only required method; the
/// other four callbacks have default no-op implementations.
pub trait ReadListener<T> {
    /// Called when row conversion or processing fails.
    ///
    /// Mirrors Java `onException(Exception, AnalysisContext) throws Exception`,
    /// where the exception is mapped to [`ErrorAction`].
    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        ErrorAction::Stop
    }

    /// Called for a resolved header row. (Java `invokeHead(Map<Integer, ReadCellData<?>>, AnalysisContext)`)
    ///
    /// # Errors
    ///
    /// Returns an error to stop the read operation.
    fn invoke_head(
        &mut self,
        _head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        Ok(())
    }

    /// Called once for every successfully converted row. (Java `invoke(T, AnalysisContext)`)
    ///
    /// # Errors
    ///
    /// Returns an error to stop the read operation.
    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()>;

    /// Called when enabled comment, hyperlink, or merge metadata is encountered.
    /// (Java `extra(CellExtra, AnalysisContext)`)
    ///
    /// # Errors
    ///
    /// Returns an error to route through [`Self::on_exception`].
    fn extra(&mut self, _extra: &CellExtra, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }

    /// Called after a sheet has been analysed. (Java `doAfterAllAnalysed(AnalysisContext)`)
    ///
    /// # Errors
    ///
    /// Returns an error when final listener work fails.
    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }

    /// Allows a listener to stop before the next row. (Java `hasNext(AnalysisContext)`)
    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        true
    }
}

/// Dispatches every read callback to two listeners in registration order.
///
/// Java stores a list of custom `ReadListener`s on `ReadBasicParameter`.
/// Rust models the same ordered fan-out as a nested, statically typed listener
/// so registering another listener does not require runtime type erasure.
pub struct CompositeReadListener<T, First, Second> {
    first: First,
    second: Second,
    marker: std::marker::PhantomData<fn() -> T>,
}

impl<T, First, Second> CompositeReadListener<T, First, Second> {
    /// Creates an ordered pair where `first` is invoked before `second`.
    #[must_use]
    pub const fn new(first: First, second: Second) -> Self {
        Self {
            first,
            second,
            marker: std::marker::PhantomData,
        }
    }

    /// Returns both listeners after a read completes.
    #[must_use]
    pub fn into_inner(self) -> (First, Second) {
        (self.first, self.second)
    }
}

/// Ordered, dynamically sized Java-style custom listener list.
///
/// This is used by compatibility builders whose listener count is only known
/// at runtime. Rows are cloned because Rust listeners own their argument,
/// while Java listeners receive the same object reference.
pub struct ReadListenerList<T> {
    listeners: Vec<Box<dyn ReadListener<T>>>,
}

impl<T> Default for ReadListenerList<T> {
    fn default() -> Self {
        Self {
            listeners: Vec::new(),
        }
    }
}

impl<T> ReadListenerList<T> {
    /// Creates a list containing its first listener.
    #[must_use]
    pub fn new(listener: impl ReadListener<T> + 'static) -> Self {
        Self {
            listeners: vec![Box::new(listener)],
        }
    }

    /// Appends a listener in Java registration order.
    pub fn push(&mut self, listener: impl ReadListener<T> + 'static) {
        self.listeners.push(Box::new(listener));
    }

    /// Appends an already boxed listener.
    pub fn push_boxed(&mut self, listener: Box<dyn ReadListener<T>>) {
        self.listeners.push(listener);
    }

    /// Returns the registered listener count.
    #[must_use]
    pub fn len(&self) -> usize {
        self.listeners.len()
    }

    /// Returns whether no listeners are registered.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.listeners.is_empty()
    }
}

impl<T> ReadListener<T> for ReadListenerList<T>
where
    T: Clone,
{
    fn on_exception(&mut self, error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        self.listeners
            .iter_mut()
            .map(|listener| listener.on_exception(error, context))
            .fold(ErrorAction::Continue, strongest_error_action)
    }

    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        for listener in &mut self.listeners {
            listener.invoke_head(head, context)?;
        }
        Ok(())
    }

    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        for listener in &mut self.listeners {
            listener.invoke(data.clone(), context)?;
        }
        Ok(())
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        for listener in &mut self.listeners {
            listener.extra(extra, context)?;
        }
        Ok(())
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        for listener in &mut self.listeners {
            listener.do_after_all_analysed(context)?;
        }
        Ok(())
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        let mut has_next = true;
        for listener in &mut self.listeners {
            has_next &= listener.has_next(context);
        }
        has_next
    }
}

const fn strongest_error_action(left: ErrorAction, right: ErrorAction) -> ErrorAction {
    match (left, right) {
        (ErrorAction::Stop, _) | (_, ErrorAction::Stop) => ErrorAction::Stop,
        (ErrorAction::SkipRow, _) | (_, ErrorAction::SkipRow) => ErrorAction::SkipRow,
        _ => ErrorAction::Continue,
    }
}

impl<T, First, Second> ReadListener<T> for CompositeReadListener<T, First, Second>
where
    T: Clone,
    First: ReadListener<T>,
    Second: ReadListener<T>,
{
    fn on_exception(&mut self, error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        let first_action = self.first.on_exception(error, context);
        let second_action = self.second.on_exception(error, context);
        strongest_error_action(first_action, second_action)
    }

    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        self.first.invoke_head(head, context)?;
        self.second.invoke_head(head, context)
    }

    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        self.first.invoke(data.clone(), context)?;
        self.second.invoke(data, context)
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        self.first.extra(extra, context)?;
        self.second.extra(extra, context)
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.first.do_after_all_analysed(context)?;
        self.second.do_after_all_analysed(context)
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        let first_has_next = self.first.has_next(context);
        let second_has_next = self.second.has_next(context);
        first_has_next && second_has_next
    }
}

impl<T, L: ReadListener<T> + ?Sized> ReadListener<T> for Box<L> {
    fn on_exception(&mut self, error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        (**self).on_exception(error, context)
    }

    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        (**self).invoke_head(head, context)
    }

    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        (**self).invoke(data, context)
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        (**self).extra(extra, context)
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        (**self).do_after_all_analysed(context)
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        (**self).has_next(context)
    }
}

impl<T, L: ReadListener<T> + ?Sized> ReadListener<T> for &mut L {
    fn on_exception(&mut self, error: &ExcelError, context: &AnalysisContext) -> ErrorAction {
        (**self).on_exception(error, context)
    }

    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        context: &AnalysisContext,
    ) -> Result<()> {
        (**self).invoke_head(head, context)
    }

    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        (**self).invoke(data, context)
    }

    fn extra(&mut self, extra: &CellExtra, context: &AnalysisContext) -> Result<()> {
        (**self).extra(extra, context)
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        (**self).do_after_all_analysed(context)
    }

    fn has_next(&mut self, context: &AnalysisContext) -> bool {
        (**self).has_next(context)
    }
}
