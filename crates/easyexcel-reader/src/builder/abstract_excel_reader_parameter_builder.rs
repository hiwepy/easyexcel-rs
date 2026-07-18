//! Mirrors Java `com.alibaba.excel.read.builder.AbstractExcelReaderParameterBuilder`.

/// Mirrors Java `AbstractExcelReaderParameterBuilder<T, C>`.
pub trait AbstractExcelReaderParameterBuilder {
    /// Sets the head row count. (Java `headRowNumber(Integer)`)
    fn head_row_number(&mut self, _head_row_number: i32) -> &mut Self
    where
        Self: Sized,
    { self }

    /// Sets the scientific format flag. (Java `useScientificFormat(Boolean)`)
    fn use_scientific_format(&mut self, _enabled: bool) -> &mut Self
    where
        Self: Sized,
    { self }

    /// Appends a read listener. (Java `registerReadListener(ReadListener)`)
    fn register_read_listener<T>(
        &mut self,
        _listener: Box<dyn easyexcel_core::ReadListener<T>>,
    ) -> &mut Self
    where
        Self: Sized,
    { self }
}
