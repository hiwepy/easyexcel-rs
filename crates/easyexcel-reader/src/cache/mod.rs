//! Mirrors Java `com.alibaba.excel.cache.*` and `cache.selector.*`.
//!
//! ## Java ↔ Rust mapping
//!
//! | Java | Rust | Notes |
//! |------|------|-------|
//! | `MapCache` | [`MapCache`] | In-memory `HashMap`-style backend |
//! | `Ehcache` | [`Ehcache`] | Disk spill via `tempfile`; **not** JVM `PersistentCacheManager` |
//! | `XlsCache` | [`XlsCache`] | Pre-built SST table for BIFF reads |
//! | `SimpleReadCacheSelector` | [`SimpleReadCacheSelector`] | 5 MB (`5_000_000` byte) Auto boundary |
//! | `EternalReadCacheSelector` | [`EternalReadCacheSelector`] | Pins Memory or Disk regardless of size |
//! | `ReadCache` | [`ReadCache`] | Shared-string put/get contract |
//!
//! XLSX SAX uses [`crate::read_cache::ReadCacheMode`] (`Auto` / `Memory` / `Disk`) wired
//! through [`ReadOptions::read_cache`] and optional [`ReadOptions::read_cache_selector`].
//! Legacy XLS reads use calamine and do not consult these selectors.

mod ehcache;
mod map_cache;
mod read_cache;
pub mod selector;
mod xls_cache;

pub use ehcache::Ehcache;
pub use map_cache::MapCache;
pub use read_cache::ReadCache;
pub use selector::{
    EternalReadCacheSelector, ReadCacheSelector, SimpleReadCacheSelector,
};
pub use xls_cache::XlsCache;

pub use read_cache::{new_disk_cache, new_map_cache, resolve_read_cache_mode};

#[cfg(test)]
mod tests;
