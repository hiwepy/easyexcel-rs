//! Mirrors Java `com.alibaba.excel.context.AnalysisContext` (the public listener
//! surface — `AnalysisContextImpl` carries additional mutable state).

use std::any::Any;

use crate::custom_read_object::CustomReadObject;
use crate::excel_error::ExcelError;

/// Read callback context equivalent to Java `AnalysisContext`.
///
/// Java exposes 14 methods plus several `@Deprecated` accessors. Rust keeps
/// only the methods actually consumed by `ReadListener` callbacks; legacy
/// getters (`getExcelType`, `getInputStream`, `getCurrentRowNum`, etc.) are
/// replaced by fields carried elsewhere in the read pipeline.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AnalysisContext {
    sheet_name: String,
    sheet_no: usize,
    row_index: u32,
    batch_index: usize,
    custom_object: Option<CustomReadObject>,
}

impl AnalysisContext {
    /// Creates a context. (Java `AnalysisContextImpl(ReadWorkbook, ExcelTypeEnum)` initial state)
    #[must_use]
    pub fn new(sheet_name: impl Into<String>, sheet_no: usize, row_index: u32) -> Self {
        Self {
            sheet_name: sheet_name.into(),
            sheet_no,
            row_index,
            batch_index: 0,
            custom_object: None,
        }
    }

    /// Returns the sheet name. (Java `AnalysisContext.readSheetHolder().getSheetName()`)
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Returns the zero-based sheet index. (Java `AnalysisContext.readSheetHolder().getSheetNo()`)
    #[must_use]
    pub const fn sheet_no(&self) -> usize {
        self.sheet_no
    }

    /// Returns the zero-based physical row index. (Java `getCurrentRowNum()`)
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the zero-based callback batch index.
    /// Rust extension tracking the page index in `PageReadListener`.
    #[must_use]
    pub const fn batch_index(&self) -> usize {
        self.batch_index
    }

    /// Returns the configured custom read object, if any.
    #[must_use]
    pub const fn custom_object(&self) -> Option<&CustomReadObject> {
        self.custom_object.as_ref()
    }

    /// Returns the custom read object when its concrete type matches `T`.
    /// Mirrors `(T) AnalysisContext.getCustom()` after an explicit cast.
    #[must_use]
    pub fn custom<T: Any>(&self) -> Option<&T> {
        self.custom_object.as_ref()?.downcast_ref()
    }

    /// Returns a context carrying the supplied custom read object.
    #[must_use]
    pub fn with_custom_object(mut self, custom_object: Option<CustomReadObject>) -> Self {
        self.custom_object = custom_object;
        self
    }

    /// Returns a copy with a different batch index.
    #[must_use]
    pub fn with_batch_index(&self, batch_index: usize) -> Self {
        let mut context = self.clone();
        context.batch_index = batch_index;
        context
    }
}

/// Action selected by a listener after a row error.
///
/// Mirrors Java `ReadListener.onException(...)` semantics:
/// * `Continue` ⇒ Java's `onException` returns without throwing.
/// * `SkipRow` ⇒ Rust extension for batch pagination.
/// * `Stop` ⇒ Java's `onException` throws `ExcelAnalysisException`.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub enum ErrorAction {
    /// Continue with the next row.
    Continue,
    /// Skip the failed row and continue.
    SkipRow,
    /// Stop and return the error. (default — matches Java's throw-exception behaviour)
    #[default]
    Stop,
}

/// Result alias used across the `easyexcel` crates.
pub type Result<T> = std::result::Result<T, ExcelError>;
