//! Mirrors Java `com.alibaba.excel.event.AbstractIgnoreExceptionReadListener`.

use std::collections::HashMap;

use crate::analysis_context::AnalysisContext;
use crate::cell_extra::CellExtra;
use crate::read_listener::ReadListener;

/// Rust port of Java `AbstractIgnoreExceptionReadListener<T>`.
///
/// Java is `@Deprecated`. The class overrides `onException`,
/// `extra`, and `hasNext` with no-op defaults. Rust mirrors the
/// same behavior through the trait's default implementations.
#[deprecated(note = "Use `ReadListener` directly")]
pub trait AbstractIgnoreExceptionReadListener<T>: ReadListener<T> {
    /// Default no-op exception handler. (Java `onException` empty body)
    fn on_exception_silent(
        &mut self,
        _exception: &crate::excel_error::ExcelError,
        _context: &AnalysisContext,
    ) {
    }

    /// Default no-op extra handler. (Java `extra` empty body)
    fn extra_silent(&mut self, _extra: &CellExtra, _context: &AnalysisContext) {}

    /// Default `hasNext` returning `true`. (Java `hasNext` returning `true`)
    fn has_next_silent(&mut self, _context: &AnalysisContext) -> bool {
        true
    }
}

// Suppress unused import.
#[allow(dead_code)]
fn _import_marker(_: HashMap<usize, String>) {}
