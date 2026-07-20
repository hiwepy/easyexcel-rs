//! Mirrors Java `com.alibaba.excel.cache.Ehcache`.
//!
//! Java binds two org.ehcache managers: a persistent disk manager
//! (`FILE_CACHE_MANAGER`, 20 GB pool) and a heap active cache
//! (`ACTIVE_CACHE_MANAGER`, sized by `maxCacheActivateBatchCount` entries or
//! deprecated `maxCacheActivateSize` MB). Strings are batched in groups of
//! [`BATCH_COUNT`] before spilling to disk.
//!
//! Rust keeps the same [`ReadCache`] surface but implements spill with
//! [`crate::read_cache::ConcurrentDiskCache`] and a process-local `tempfile`.
//! There is **no** JNI / JVM `PersistentCacheManager`. Behaviour that depends
//! on cross-process persistence or Ehcache eviction policies is out of scope;
//! sequential put/get/destroy semantics are covered by unit tests.

use easyexcel_core::Result;

use super::read_cache::{new_disk_cache, ReadCache, SharedStringCacheAdapter};
use crate::read_cache::SharedStringCache;

/// Batch count used by Java `Ehcache.BATCH_COUNT`.
pub const BATCH_COUNT: usize = 100;

/// Default active batch count used by Java `SimpleReadCacheSelector`.
pub const DEFAULT_MAX_EHCACHE_ACTIVATE_BATCH_COUNT: i32 = 20;

/// Disk-backed shared-string cache matching Java `Ehcache`.
///
/// Mirrors Java `com.alibaba.excel.cache.Ehcache`.
///
/// Use [`ReadCacheMode::Disk`](crate::ReadCacheMode::Disk) or
/// [`EternalReadCacheSelector::ehcache`] at the workbook level; this type exists
/// for API parity and direct [`ReadCache`] tests.
pub struct Ehcache {
    adapter: SharedStringCacheAdapter,
}

impl Ehcache {
    /// Creates a disk-backed cache with Java default batch sizing.
    ///
    /// Mirrors Java `new Ehcache(null, maxCacheActivateBatchCount)`.
    ///
    /// The `max_cache_activate_batch_count` argument is accepted for signature
    /// parity; Rust's disk backend does not replicate Java's heap active-cache
    /// tier sizing.
    ///
    /// # Errors
    ///
    /// Returns an I/O error when the temporary cache file cannot be created.
    pub fn new(max_cache_activate_batch_count: Option<i32>) -> Result<Self> {
        let _ = max_cache_activate_batch_count
            .unwrap_or(DEFAULT_MAX_EHCACHE_ACTIVATE_BATCH_COUNT);
        Ok(Self::from_backend(new_disk_cache()?))
    }

    /// Creates a cache with the deprecated Java `maxCacheActivateSize` MB knob.
    ///
    /// Mirrors Java `new Ehcache(maxCacheActivateSize)`.
    ///
    /// # Errors
    ///
    /// Returns an I/O error when the temporary cache file cannot be created.
    pub fn with_max_cache_activate_size_mb(max_cache_activate_size_mb: Option<i32>) -> Result<Self> {
        let _ = max_cache_activate_size_mb;
        Self::new(None)
    }

    /// Wraps an existing shared-string backend.
    #[must_use]
    pub fn from_backend(backend: Box<dyn SharedStringCache>) -> Self {
        Self {
            adapter: SharedStringCacheAdapter::new(backend),
        }
    }
}

impl ReadCache for Ehcache {
    fn put(&mut self, value: String) -> Result<()> {
        self.adapter.put(value)
    }

    fn get(&self, key: Option<usize>) -> Result<Option<String>> {
        self.adapter.get(key)
    }

    fn put_finished(&mut self) -> Result<()> {
        self.adapter.put_finished()
    }
}
