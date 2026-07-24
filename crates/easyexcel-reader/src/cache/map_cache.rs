//! Mirrors Java `com.alibaba.excel.cache.MapCache`.
//!
//! Java stores every shared string in a single in-memory `ArrayList`.
//! Rust uses [`crate::read_cache::MemorySharedStringCache`] via
//! [`super::read_cache::new_map_cache`]. Selected automatically when
//! [`SimpleReadCacheSelector`] sees `sharedStrings.xml` smaller than 5 MB, or
//! when [`EternalReadCacheSelector::map_cache`] pins memory mode.

use easyexcel_core::Result;

use super::read_cache::{ReadCache, SharedStringCacheAdapter, new_map_cache};
use crate::read_cache::SharedStringCache;

/// In-memory shared-string cache matching Java `MapCache`.
///
/// Mirrors Java `com.alibaba.excel.cache.MapCache`.
pub struct MapCache {
    adapter: SharedStringCacheAdapter,
}

impl MapCache {
    /// Creates an empty map-backed cache. (Java `new MapCache()`)
    #[must_use]
    pub fn new() -> Self {
        Self::from_backend(new_map_cache())
    }

    /// Wraps an existing shared-string backend.
    #[must_use]
    pub fn from_backend(backend: Box<dyn SharedStringCache>) -> Self {
        Self {
            adapter: SharedStringCacheAdapter::new(backend),
        }
    }
}

impl Default for MapCache {
    fn default() -> Self {
        Self::new()
    }
}

impl ReadCache for MapCache {
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
