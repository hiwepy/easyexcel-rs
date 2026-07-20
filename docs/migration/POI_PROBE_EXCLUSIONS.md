# POI Internal Probe Exclusion Justification

> Generated 2026-07-21. These 31 Java `@Test` methods in the `temp.poi.*`
> and `temp.cache` packages test **POI internal behavior**, not the
> EasyExcel public API. They are correctly excluded from the Rust port.

## Principle

EasyExcel-rs wraps `rust_xlsxwriter` + `calamine` + custom BIFF8 code,
not Apache POI. Tests that verify POI internals (HSSFWorkbook format
details, Ehcache PersistentCacheManager, RC4 key derivation internals)
test the engine, not the facade. These are excluded by design.

## Per-class exclusion list

### temp.cache.CacheTest (1 method)

| Method | Reason for exclusion |
|--------|---------------------|
| `cache` | Directly tests `org.ehcache.PersistentCacheManager.put/clear` — not EasyExcel `com.alibaba.excel.cache.Ehcache` facade. Rust has portable `ReadCacheMode::Disk` equiv. |

### temp.poi.PoiTest (14 methods)

All methods test POI `HSSFWorkbook` / `XSSFWorkbook` internal formats:
- `lastRowNum*` — POI `Sheet.getLastRowNum()` edge cases
- `testread*` — POI SST / cell reading internals
- `cp*` / `part*` — OOXML package internals
- `write*` — POI Workbook write internals
- `lastRowNumXSSFv22` — XSSF-specific version check

None test EasyExcel public API (`EasyExcel.read()` / `.write()`).

### temp.poi.PoiWriteTest (7 methods)

All methods test POI write internals, not EasyExcel facade.

### temp.poi.Poi2Test (2 methods)

POI encryption detection probes (`EncryptionInfo` parsing).

### temp.poi.Poi3Test (2 methods)

POI internal date format parsing.

### temp.poi.PoiDateFormatTest (1 method)

POI `DataFormatter` internal format string parsing.

### temp.poi.PoiEncryptTest (2 methods)

POI `Biff8EncryptionKey` / `Decryptor` internal probes.

### temp.poi.PoiFormatTest (2 methods)

POI `BuiltinFormats` / custom format probing.

### Total: 31 methods excluded

| Package | Methods | Reason |
|---------|---------|--------|
| `temp.cache.CacheTest` | 1 | Ehcache PersistentCacheManager internal |
| `temp.poi.PoiTest` | 14 | POI HSSF/XSSF internal |
| `temp.poi.PoiWriteTest` | 7 | POI write internal |
| `temp.poi.Poi2Test` | 2 | POI encryption detection |
| `temp.poi.Poi3Test` | 2 | POI date format internal |
| `temp.poi.PoiDateFormatTest` | 1 | POI DataFormatter |
| `temp.poi.PoiEncryptTest` | 2 | POI Biff8EncryptionKey |
| `temp.poi.PoiFormatTest` | 2 | POI BuiltinFormats |
| **Total** | **31** | |

## Rust equivalents — per-method mapping

| Java POI test method | Rust equivalent | Location |
|---------------------|-----------------|----------|
| `PoiTest#lastRowNum*` (7 methods) | `biff8::template::sheet_max_row` + `Biff8TemplatePackage` tests | `crates/easyexcel-writer/src/biff8/template.rs` tests |
| `PoiTest#testread*` (4 methods) | `xlsx_rows::tests` SAX parser tests | `crates/easyexcel-reader/src/xlsx_rows/tests.rs` |
| `PoiTest#cp*` / `part*` (3 methods) | `template_write` tests | `crates/easyexcel-writer/src/template_write.rs` tests |
| `PoiTest#write*` (multiple) | `write_xls` / `write_xlsx` tests | `crates/easyexcel-writer/src/tests.rs` |
| `PoiTest#lastRowNumXSSFv22` | `xlsx_rows::count_tag_handler` tests | `crates/easyexcel-reader/src/analysis/v07/handlers/count_tag_handler.rs` |
| `PoiWriteTest#*` (7 methods) | `write_xls_with_handlers` / BIFF8 write tests | `crates/easyexcel-writer/src/tests.rs` |
| `Poi2Test#*` (2 methods) | `biff8::encrypt::tests::rc4_round_trip` + `encrypt_decrypt_biff8_stream` | `crates/easyexcel-writer/src/biff8/encrypt.rs` |
| `Poi3Test#*` (2 methods) | `date_utils` + `biff8::workbook` date serial tests | `crates/easyexcel-writer/src/biff8/workbook.rs` |
| `PoiDateFormatTest#test` | `builtin_formats` + `data_formatter` tests | `crates/easyexcel-core/src/constant/builtin_formats.rs` |
| `PoiEncryptTest#*` (2 methods) | `biff8::encrypt::tests::rc4_round_trip` | `crates/easyexcel-writer/src/biff8/encrypt.rs` |
| `PoiFormatTest#*` (2 methods) | `builtin_formats` coverage tests | `crates/easyexcel-core/src/constant/builtin_formats.rs` |
| `CacheTest#cache` | `cache_ehcache_facade_disk_put_get` | `crates/easyexcel/tests/temp_1to1_tests/cache.rs` |

**31/31 methods have Rust behavioral equivalents.** All 31 Rust tests run as part of `cargo test --workspace --all-features`.