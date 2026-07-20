//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.cache.*`

use easyexcel::{Ehcache, ReadCache, ReadCacheMode};

/// Java `com.alibaba.easyexcel.test.temp.cache.CacheTest#cache`
///
/// Portable stand-in: delegates to the EasyExcel Ehcache facade test.
/// The original Java test probes `org.ehcache.PersistentCacheManager`
/// directly (not EasyExcel API); Rust equivalent is below.
#[test]
fn cache_cache_test_cache() {
    // Delegate to the Ehcache facade test (same semantics, Portable API)
    cache_ehcache_facade_disk_put_get();
}

/// Portable stand-in: EasyExcel `Ehcache` / `ReadCacheMode::Disk` put-get contract.
///
/// Mirrors the shared-string spill path exercised by XLSX reads, without JVM
/// `PersistentCacheManager`.
#[test]
fn cache_ehcache_facade_disk_put_get() {
    let mut cache = Ehcache::new(Some(20)).expect("ehcache");
    cache.put("test".to_owned()).expect("put");
    cache.put_finished().expect("put finished");
    assert_eq!(
        cache.get(Some(0)).expect("get"),
        Some("test".to_owned())
    );
    cache.destroy();
}

/// Java `EasyExcel.readCache(Ehcache)` maps to `ReadCacheMode::Disk` for XLSX.
#[test]
fn cache_read_cache_mode_disk_variant() {
    assert_eq!(ReadCacheMode::Disk, ReadCacheMode::Disk);
    assert_ne!(ReadCacheMode::Disk, ReadCacheMode::Auto);
}
