//! Cache facade tests.

use super::{
    Ehcache, EternalReadCacheSelector, MapCache, ReadCache, ReadCacheSelector,
    SimpleReadCacheSelector, XlsCache,
};
use crate::ReadCacheMode;

#[test]
fn map_cache_stores_and_retrieves_values() {
    let mut cache = MapCache::new();
    cache.put("alpha".to_owned()).expect("put");
    cache.put("beta".to_owned()).expect("put");
    cache.put_finished().expect("put finished");
    assert_eq!(
        cache.get(Some(1)).expect("get"),
        Some("beta".to_owned())
    );
}

#[test]
fn xls_cache_reads_preloaded_sst_values() {
    let cache = XlsCache::new(vec!["one".to_owned(), "two".to_owned()]);
    assert_eq!(cache.len(), 2);
    assert_eq!(
        cache.get(Some(0)).expect("get"),
        Some("one".to_owned())
    );
    assert!(cache.get(Some(99)).expect("get").is_none());
}

#[test]
fn simple_selector_matches_java_five_megabyte_threshold() {
    let selector = SimpleReadCacheSelector::new();
    assert_eq!(
        selector.select_mode(4_999_999),
        ReadCacheMode::Memory
    );
    assert_eq!(selector.select_mode(5_000_000), ReadCacheMode::Disk);
}

#[test]
fn simple_selector_custom_mb_threshold() {
    let selector = SimpleReadCacheSelector::with_max_use_map_cache_size_mb(1);
    assert_eq!(selector.select_mode(999_999), ReadCacheMode::Memory);
    assert_eq!(selector.select_mode(1_000_000), ReadCacheMode::Disk);
}

#[test]
fn eternal_selector_pins_backend_mode() {
    let selector = EternalReadCacheSelector::map_cache();
    assert_eq!(selector.select_mode(9_999_999), ReadCacheMode::Memory);
    let disk = EternalReadCacheSelector::ehcache();
    assert_eq!(disk.select_mode(0), ReadCacheMode::Disk);
}

#[test]
fn ehcache_round_trips_values_through_disk_backend() {
    let mut cache = Ehcache::new(Some(20)).expect("ehcache");
    cache.put("disk-value".to_owned()).expect("put");
    cache.put_finished().expect("put finished");
    assert_eq!(
        cache.get(Some(0)).expect("get"),
        Some("disk-value".to_owned())
    );
    cache.destroy();
}

#[test]
fn ehcache_deprecated_size_constructor_is_accepted() {
    let mut cache = Ehcache::with_max_cache_activate_size_mb(Some(20)).expect("ehcache");
    cache.put("legacy-knob".to_owned()).expect("put");
    cache.put_finished().expect("put finished");
    assert_eq!(
        cache.get(Some(0)).expect("get"),
        Some("legacy-knob".to_owned())
    );
}
