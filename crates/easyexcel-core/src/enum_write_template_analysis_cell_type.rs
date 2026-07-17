//! Mirrors Java `com.alibaba.excel.enums.WriteTemplateAnalysisCellTypeEnum`.

/// Cell kind discovered while analysing a template placeholder.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WriteTemplateAnalysisCellType {
    /// Common placeholder such as `{key}`.
    Common,
    /// Collection placeholder such as `{name.field}`.
    Collection,
}
