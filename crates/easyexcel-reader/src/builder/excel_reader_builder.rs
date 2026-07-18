//! Mirrors Java `com.alibaba.excel.read.builder.ExcelReaderBuilder`.

/// Mirrors Java `ExcelReaderBuilder extends AbstractExcelReaderParameterBuilder`.
#[derive(Debug, Clone, Default)]
pub struct ExcelReaderBuilder {
    /// Mirrors `ExcelReaderBuilder.file`.
    pub file: Option<String>,
}

impl ExcelReaderBuilder {
    /// Creates a builder.
    #[must_use]
    pub fn new() -> Self { Self::default() }

    /// Sets the file path. (Java `file(File)`)
    pub fn file(mut self, path: impl Into<String>) -> Self {
        self.file = Some(path.into());
        self
    }
}
