//! Mirrors Java `com.alibaba.excel.read.builder.AbstractExcelReaderParameterBuilder`.

/// Mirrors Java `AbstractExcelReaderParameterBuilder<T, C>`.
pub trait AbstractExcelReaderParameterBuilder<T> {
    /// Sets the head row count. (Java `headRowNumber(Integer)`)
    fn head_row_number(&mut self, head_row_number: i32) -> &mut Self;

    /// Sets the scientific format flag. (Java `useScientificFormat(Boolean)`)
    fn use_scientific_format(&mut self, enabled: bool) -> &mut Self;

    /// Appends a read listener. (Java `registerReadListener(ReadListener)`)
    fn register_read_listener(
        &mut self,
        listener: Box<dyn easyexcel_core::ReadListener<T>>,
    ) -> &mut Self;
}
