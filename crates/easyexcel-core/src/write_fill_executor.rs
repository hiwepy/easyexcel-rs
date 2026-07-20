//! Mirrors Java `com.alibaba.excel.write.executor.ExcelWriteFillExecutor`.
//!
//! The trait lives in `easyexcel-core` so `easyexcel-writer` can hold an optional
//! hook without depending on `easyexcel-template`, and the template crate can
//! provide the concrete fill implementation.

use std::any::Any;

use crate::{ExcelError, Result, WriteDirection};

/// Minimal fill configuration at the [`ExcelBuilder`](crate::WriteContext) surface.
///
/// Mirrors Java `com.alibaba.excel.write.metadata.fill.FillConfig` fields used by
/// `ExcelBuilderImpl.fill`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct WriteFillConfig {
    /// Whether collection fill forces a new row. (Java `FillConfig.forceNewRow`)
    pub force_new_row: bool,
    /// Collection expansion direction when supplied by the caller.
    pub direction: Option<WriteDirection>,
}

impl WriteFillConfig {
    /// Creates Java-compatible defaults (`forceNewRow = false`, vertical fill).
    #[must_use]
    pub const fn new() -> Self {
        Self {
            force_new_row: false,
            direction: None,
        }
    }
}

/// Worksheet metadata passed into template fill execution.
///
/// Mirrors Java `WriteSheet` selection inside `ExcelBuilderImpl.fill`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteFillSheet {
    /// Selected worksheet name.
    pub sheet_name: String,
    /// Optional zero-based sheet index.
    pub sheet_index: Option<usize>,
}

impl Default for WriteFillSheet {
    fn default() -> Self {
        Self {
            sheet_name: "Sheet1".to_owned(),
            sheet_index: None,
        }
    }
}

/// Hook implemented by `easyexcel-template` and wired from the `easyexcel` facade.
///
/// Mirrors Java `ExcelWriteFillExecutor.fill(Object, FillConfig)`.
pub trait WriteFillExecutor {
    /// Accumulates one scalar or collection fill against the loaded template.
    ///
    /// # Errors
    ///
    /// Returns a format error when `data` is not a supported fill payload, or a
    /// template I/O / OOXML error from the underlying engine.
    fn fill(
        &mut self,
        data: &dyn Any,
        fill_config: WriteFillConfig,
        sheet: WriteFillSheet,
    ) -> Result<()>;

    /// Persists accumulated fill results to the configured output target.
    ///
    /// Mirrors Java `WriteContext.finish(boolean onException)` for fill-only
    /// sessions.
    ///
    /// # Errors
    ///
    /// Returns an output, close, or package-format error.
    fn finish(&mut self, on_exception: bool) -> Result<()>;
}

/// Returns a descriptive error when no template stream is configured.
///
/// Mirrors Java `ExcelGenerateException("Calling the 'fill' method must use a template.")`.
#[must_use]
pub fn fill_requires_template_error() -> ExcelError {
    ExcelError::Unsupported(
        "Calling the 'fill' method must use a template.".to_owned(),
    )
}

/// Returns a descriptive error when CSV fill is requested.
///
/// Mirrors Java `ExcelGenerateException("csv does not support filling data.")`.
#[must_use]
pub fn csv_fill_unsupported_error() -> ExcelError {
    ExcelError::Unsupported("csv does not support filling data.".to_owned())
}
