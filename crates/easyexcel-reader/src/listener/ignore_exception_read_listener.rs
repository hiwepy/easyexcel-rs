//! Mirrors Java `com.alibaba.excel.read.listener.IgnoreExceptionReadListener`.

use easyexcel_core::{AnalysisContext, ReadListener};

/// Mirrors Java `IgnoreExceptionReadListener extends ReadListener<T>`.
///
/// Java overrides `onException` to swallow the error and `hasNext` to
/// return `true`. The Rust port implements the same defaults on the
/// trait.
pub trait IgnoreExceptionReadListener<T>: ReadListener<T> {
    /// Default exception handler that returns `ErrorAction::Continue`
    /// instead of the trait's `Stop` default. (Java `onException` empty body)
    fn on_exception_silent(
        &mut self,
        _error: &easyexcel_core::ExcelError,
        _context: &AnalysisContext,
    ) -> easyexcel_core::ErrorAction {
        easyexcel_core::ErrorAction::Continue
    }
}
