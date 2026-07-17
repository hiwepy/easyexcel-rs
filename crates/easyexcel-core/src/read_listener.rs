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
