//! Mirrors Java `com.alibaba.excel.enums.CacheLocationEnum`.
//!
//! Used by Java `BasicParameter.filedCacheLocation`. Rust has collapsed this
//! concept into `easyexcel_reader::ReadCacheMode`, but the enum is kept for
//! API completeness when reading Java `ReadWorkbookHolder` payloads.

/// Cache location strategy.
///
/// Rust port of Java `CacheLocationEnum`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CacheLocation {
    /// Stored in `ThreadLocal`; cleared when the read or write completes.
    ThreadLocal,
    /// Never cleared unless the application exits.
    Memory,
    /// Caching disabled.
    None,
}
