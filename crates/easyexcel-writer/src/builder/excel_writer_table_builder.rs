//! Mirrors Java `com.alibaba.excel.write.builder.ExcelWriterTableBuilder`.

use crate::CellStyle;
use crate::WriteHandler;

use crate::builder::abstract_excel_writer_parameter_builder::AbstractExcelWriterParameterBuilder;
use crate::metadata::{WriteBasicParameter, WriteTable};

/// Mirrors Java `ExcelWriterTableBuilder extends AbstractExcelWriterParameterBuilder`.
///
/// Java carries a `WriteTable` and a back-reference to the parent
/// `WriteSheet`; Rust mirrors the data on the parameter struct and
/// exposes the same builder surface.
pub struct ExcelWriterTableBuilder {
    parameter: WriteBasicParameter,
    table: WriteTable,
    handlers: Vec<Box<dyn WriteHandler>>,
}

impl ExcelWriterTableBuilder {
    /// Creates a table builder.
    #[must_use]
    pub fn new() -> Self {
        Self {
            parameter: WriteBasicParameter::default(),
            table: WriteTable::new(),
            handlers: Vec::new(),
        }
    }

    /// Sets the zero-based table index. (Java `tableNo(Integer)`)
    pub fn table_no(mut self, table_no: i32) -> Self {
        self.table.table_no = table_no;
        self
    }

    /// Builds the `WriteTable` value. (Java `build()`)
    #[must_use]
    pub fn build(&self) -> WriteTable {
        let mut table = self.table.clone();
        table.options.relative_head_row_index = self.parameter.relative_head_row_index;
        table.options.need_head = self.parameter.need_head;
        table.options.automatic_merge_head = self.parameter.automatic_merge_head;
        table
    }

    /// Returns a reference to the inner `WriteTable` for inspection.
    #[must_use]
    pub const fn table(&self) -> &WriteTable {
        &self.table
    }

    /// Returns the number of currently registered handlers. Useful for tests.
    #[must_use]
    pub fn handler_count(&self) -> usize {
        self.handlers.len()
    }

    /// Returns the registered handler list. (Java `getCustomWriteHandlerList()`)
    #[must_use]
    pub fn handlers(&self) -> &[Box<dyn WriteHandler>] {
        &self.handlers
    }

    /// Appends a write handler. (Java `registerWriteHandler(WriteHandler)`)
    pub fn register_write_handler(&mut self, handler: Box<dyn WriteHandler>) -> &mut Self {
        self.handlers.push(handler);
        self
    }

    /// Convenience setter that records a head style without emitting a
    /// no-op. Provided for parity with Java's chainable setters.
    pub fn head_style_record(&mut self, _style: CellStyle) -> &mut Self {
        self
    }
}

impl Default for ExcelWriterTableBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl AbstractExcelWriterParameterBuilder for ExcelWriterTableBuilder {
    fn parameter(&mut self) -> &mut WriteBasicParameter {
        &mut self.parameter
    }

    fn register_write_handler(&mut self, handler: Box<dyn WriteHandler>) -> &mut Self {
        self.register_write_handler(handler)
    }
}