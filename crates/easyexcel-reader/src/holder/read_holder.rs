//! Mirrors Java `com.alibaba.excel.read.metadata.holder.ReadHolder` (interface).

use easyexcel_core::AnalysisContext;

/// Mirrors Java `ReadHolder extends ConfigurationHolder`.
pub trait ReadHolder {
    /// Returns the analysis context. (Java `getAnalysisContext()`)
    fn analysis_context(&self) -> &AnalysisContext;
}
