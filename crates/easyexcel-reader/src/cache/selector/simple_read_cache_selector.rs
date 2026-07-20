//! Mirrors Java `com.alibaba.excel.cache.selector.SimpleReadCacheSelector`.
//!
//! Default workbook behaviour (`ReadCacheMode::Auto`) uses the same
//! [`DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES`] (5_000_000) boundary as Java
//! `EasyExcel`'s built-in selector: smaller `sharedStrings.xml` parts stay in
//! [`MapCache`](super::super::MapCache), larger parts spill to
//! [`Ehcache`](super::super::Ehcache) / disk.

use super::read_cache_selector::ReadCacheSelector;
use crate::read_cache::{ReadCacheMode, DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES};

/// Simple cache selector matching Java's 5 MB map-cache boundary.
///
/// Mirrors Java `com.alibaba.excel.cache.selector.SimpleReadCacheSelector`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SimpleReadCacheSelector {
    /// Maximum shared-string table size that keeps data in memory, in bytes.
    max_use_map_cache_size_bytes: u64,
    /// Deprecated Java `maxCacheActivateSize` placeholder kept for parity.
    max_cache_activate_size: Option<i32>,
    /// Maximum in-memory batch count for disk cache activation.
    max_cache_activate_batch_count: Option<i32>,
}

impl Default for SimpleReadCacheSelector {
    fn default() -> Self {
        Self {
            max_use_map_cache_size_bytes: DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES,
            max_cache_activate_size: None,
            max_cache_activate_batch_count: None,
        }
    }
}

impl SimpleReadCacheSelector {
    /// Creates a selector with Java defaults.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Creates a selector with a custom map-cache size in megabytes.
    ///
    /// Mirrors Java `SimpleReadCacheSelector(Long maxUseMapCacheSize, Integer maxCacheActivateSize)`.
    #[must_use]
    pub fn with_max_use_map_cache_size_mb(max_use_map_cache_size_mb: u64) -> Self {
        Self {
            max_use_map_cache_size_bytes: max_use_map_cache_size_mb.saturating_mul(1_000_000),
            ..Self::default()
        }
    }

    /// Sets the map-cache threshold in megabytes. (Java `setMaxUseMapCacheSize`)
    #[must_use]
    pub fn max_use_map_cache_size_mb(mut self, megabytes: u64) -> Self {
        self.max_use_map_cache_size_bytes = megabytes.saturating_mul(1_000_000);
        self
    }

    /// Sets the deprecated activate-size knob. (Java `setMaxCacheActivateSize`)
    #[must_use]
    pub const fn max_cache_activate_size(mut self, size: Option<i32>) -> Self {
        self.max_cache_activate_size = size;
        self
    }

    /// Sets the in-memory batch count. (Java `setMaxCacheActivateBatchCount`)
    #[must_use]
    pub const fn max_cache_activate_batch_count(mut self, count: Option<i32>) -> Self {
        self.max_cache_activate_batch_count = count;
        self
    }

    /// Returns the configured map-cache threshold in bytes.
    #[must_use]
    pub const fn max_use_map_cache_size_bytes(&self) -> u64 {
        self.max_use_map_cache_size_bytes
    }

    /// Returns the deprecated activate-size knob. (Java `getMaxCacheActivateSize()`)
    #[must_use]
    pub const fn get_max_cache_activate_size(&self) -> Option<i32> {
        self.max_cache_activate_size
    }

    /// Returns the configured batch-count knob. (Java `getMaxCacheActivateBatchCount()`)
    #[must_use]
    pub const fn get_max_cache_activate_batch_count(&self) -> Option<i32> {
        self.max_cache_activate_batch_count
    }
}

impl ReadCacheSelector for SimpleReadCacheSelector {
    fn select_mode(&self, shared_strings_xml_size: u64) -> ReadCacheMode {
        if shared_strings_xml_size < self.max_use_map_cache_size_bytes {
            ReadCacheMode::Memory
        } else {
            ReadCacheMode::Disk
        }
    }
}
