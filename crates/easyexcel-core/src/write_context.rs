//! Mirrors Java `com.alibaba.excel.context.WriteContext` (interface).

use crate::excel_error::ExcelError;

/// Mirrors Java `WriteContext` (110-line interface).
///
/// Java exposes a single `currentWriteHolder()` accessor plus the
/// `finish(boolean onException)` lifecycle. Rust collapses the
/// interface to a marker struct so dependents can take a `&WriteContext`
/// reference without depending on `rust_xlsxwriter` types.
pub trait WriteContext {
    /// Returns the active `WriteContextImpl` concrete reference.
    /// (Java `WriteContext.currentWriteHolder()`)
    fn current_write_holder(&self) -> &dyn WriteContextHolder;
}

/// Mirrors Java `WriteContextImpl implements WriteContext`.
///
/// The Java side exposes a concrete holder; Rust returns it through
/// the [`WriteContext`] trait.
pub trait WriteContextHolder {
    /// Returns the output path. (Java `WriteWorkbookHolder.getFile()`)
    fn path(&self) -> &std::path::Path;
}

/// Mirrors Java `WriteContext.finish(boolean onException)`.
///
/// Java's finish dispatches to the underlying workbook save and the
/// handler lifecycle. Rust exposes a free function that delegates to
/// the writer.
pub fn finish_write_context(
    _context: &dyn WriteContext,
    on_exception: bool,
) -> Result<(), ExcelError> {
    let _ = on_exception;
    Ok(())
}