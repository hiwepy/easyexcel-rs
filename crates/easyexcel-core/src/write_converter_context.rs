//! Mirrors Java `com.alibaba.excel.converters.WriteConverterContext`.

use crate::convert_context::ConvertContext;
use crate::excel_column::ExcelColumn;

/// Context supplied to a custom Rust-to-cell converter.
///
/// Mirrors Java `WriteConverterContext<T>(value, contentProperty,
/// writeContext)`. Rust drops the heavy `WriteContext` reference and uses
/// the lightweight `ConvertContext`.
#[derive(Debug, Clone, Copy)]
pub struct WriteConverterContext<'a, T> {
    value: &'a T,
    column: &'a ExcelColumn,
    context: &'a ConvertContext,
}

impl<'a, T> WriteConverterContext<'a, T> {
    /// Creates a write conversion context. (Java `@AllArgsConstructor`)
    #[must_use]
    pub const fn new(value: &'a T, column: &'a ExcelColumn, context: &'a ConvertContext) -> Self {
        Self {
            value,
            column,
            context,
        }
    }

    /// Returns the Rust field value. (Java `getValue()`)
    #[must_use]
    pub const fn value(&self) -> &'a T {
        self.value
    }

    /// Returns the field's static column metadata. (Java `getContentProperty()`)
    #[must_use]
    pub const fn column(&self) -> &'a ExcelColumn {
        self.column
    }

    /// Returns the target row, column, field, and format information. (Java `getWriteContext()`)
    #[must_use]
    pub const fn convert_context(&self) -> &'a ConvertContext {
        self.context
    }
}
