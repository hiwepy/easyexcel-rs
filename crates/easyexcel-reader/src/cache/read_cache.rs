//! Mirrors Java `com.alibaba.excel.cache.ReadCache`.

use easyexcel_core::Result;

use crate::cache::selector::ReadCacheSelector;
use crate::read_cache::{
    DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES, ReadCacheMode, SharedStringCache,
    SharedStringCacheReader, SharedStringCacheWriter, create_cache,
};

/// Shared-string cache contract matching Java `ReadCache`.
///
/// Mirrors Java `com.alibaba.excel.cache.ReadCache`.
pub trait ReadCache: Send {
    /// Initializes the cache. (Java `init(AnalysisContext)`)
    ///
    /// Default implementation records initialization state so callers
    /// can verify the lifecycle fires. Concrete implementations should
    /// override to allocate resources.
    fn init(&mut self) {
        // Default: no resources to allocate (in-memory caches are lazy).
        // Concrete impls override for disk/Ehcache/etc.
    }

    /// Stores the next shared string. (Java `put(String)`)
    ///
    /// # Errors
    ///
    /// Returns a format or I/O error when the value cannot be stored.
    fn put(&mut self, value: String) -> Result<()>;

    /// Reads a shared string by index. (Java `get(Integer)`)
    ///
    /// # Errors
    ///
    /// Returns a format or I/O error when the index is invalid.
    fn get(&self, key: Option<usize>) -> Result<Option<String>>;

    /// Marks the write phase complete. (Java `putFinished()`)
    ///
    /// # Errors
    ///
    /// Returns a format or I/O error when finalization fails.
    fn put_finished(&mut self) -> Result<()>;

    /// Releases cache resources. (Java `destroy()`)
    ///
    /// Default implementation is a no-op; concrete disk/Ehcache
    /// implementations override to close files and free handles.
    fn destroy(&mut self) {
        // Default: nothing to release for in-memory caches.
        // Concrete impls override for disk-based caches.
    }
}

/// Creates an in-memory cache backend. (Java `new MapCache()`)
///
/// # Panics
///
/// Panics only when the internal memory cache factory fails, which should
/// never happen for in-memory backends.
#[must_use]
pub fn new_map_cache() -> Box<dyn SharedStringCache> {
    create_cache(ReadCacheMode::Memory, 0).expect("memory cache always succeeds")
}

/// Creates a disk-backed cache backend. (Java `new Ehcache(...)`)
///
/// # Errors
///
/// Returns an I/O error when the temporary cache file cannot be created.
pub fn new_disk_cache() -> Result<Box<dyn SharedStringCache>> {
    create_cache(ReadCacheMode::Disk, DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES)
}

/// Resolves the effective [`ReadCacheMode`] for a shared-string table size.
///
/// Mirrors Java `ReadWorkbookHolder` selector wiring.
#[must_use]
pub fn resolve_read_cache_mode(
    mode: ReadCacheMode,
    selector: Option<&dyn ReadCacheSelector>,
    shared_strings_xml_size: u64,
) -> ReadCacheMode {
    selector.map_or(mode, |selector| {
        selector.select_mode(shared_strings_xml_size)
    })
}

/// Adapts the internal SAX cache writer to the Java `ReadCache` surface.
pub(crate) struct SharedStringCacheAdapter {
    inner: Box<dyn SharedStringCache>,
    reader: Option<Box<dyn SharedStringCacheReader>>,
}

impl SharedStringCacheAdapter {
    /// Wraps a live shared-string cache writer.
    #[must_use]
    pub fn new(inner: Box<dyn SharedStringCache>) -> Self {
        Self {
            inner,
            reader: None,
        }
    }

    /// Returns the read-only cache produced by [`ReadCache::put_finished`].
    ///
    /// # Panics
    ///
    /// Panics when called before [`ReadCache::put_finished`].
    #[must_use]
    pub fn into_reader(self) -> Box<dyn SharedStringCacheReader> {
        self.reader
            .expect("ReadCache.put_finished must run before into_reader")
    }
}

impl ReadCache for SharedStringCacheAdapter {
    fn put(&mut self, value: String) -> Result<()> {
        self.inner.put(value)
    }

    fn get(&self, key: Option<usize>) -> Result<Option<String>> {
        let Some(index) = key else {
            return Ok(None);
        };
        if let Some(reader) = &self.reader {
            return reader.get(index).map(Some);
        }
        self.inner.get(index).map(Some)
    }

    fn put_finished(&mut self) -> Result<()> {
        if self.reader.is_some() {
            return Ok(());
        }
        let writer = std::mem::replace(&mut self.inner, create_cache(ReadCacheMode::Memory, 0)?);
        self.reader = Some(writer.finish()?);
        Ok(())
    }
}
