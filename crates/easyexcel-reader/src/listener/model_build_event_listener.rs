//! Mirrors Java `com.alibaba.excel.read.listener.ModelBuildEventListener`.

/// Mirrors Java `ModelBuildEventListener implements IgnoreExceptionReadListener`.
///
/// Java's listener uses reflection to populate a typed JavaBean. Rust
/// reproduces the same shape through the derive macro and `RowData`.
#[derive(Debug, Clone, Default)]
pub struct ModelBuildEventListener;
