//! Mirrors Java `com.alibaba.excel.cache.selector.ReadCacheSelector`.

use crate::read_cache::ReadCacheMode;

/// Selects the shared-string cache backend for an XLSX workbook.
///
/// Mirrors Java `com.alibaba.excel.cache.selector.ReadCacheSelector`.
///
/// Java receives the `sharedStrings.xml` package part size in bytes. Rust passes
/// the same measurement into [`select_mode`](Self::select_mode). Use
/// [`SimpleReadCacheSelector`] for the default 5 MB Auto boundary, or
/// [`EternalReadCacheSelector`] to pin Memory/Disk regardless of size.
pub trait ReadCacheSelector: Send + Sync {
    /// Selects a cache mode for the given `sharedStrings.xml` size.
    ///
    /// Mirrors Java `readCache(PackagePart sharedStringsTablePackagePart)`.
    fn select_mode(&self, shared_strings_xml_size: u64) -> ReadCacheMode;
}
