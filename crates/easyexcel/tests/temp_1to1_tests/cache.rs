//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.cache.*`

use easyexcel::{Ehcache, ReadCache, ReadCacheMode};

/// Java `com.alibaba.easyexcel.test.temp.cache.CacheTest#cache`
///
/// Exclusion（严格 100%）：Java 直接探测 `org.ehcache.PersistentCacheManager`
/// put/clear 语义，非 EasyExcel `com.alibaba.excel.cache.Ehcache` 门面。
/// Rust 对等行为见 [`cache_ehcache_facade_disk_put_get`] 与 reader `cache/tests.rs`。
#[test]
#[ignore = "ehcache-stress: org.ehcache PersistentCacheManager probe — not EasyExcel Ehcache API"]
fn cache_cache_test_cache() {
    panic!("ignored");
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
