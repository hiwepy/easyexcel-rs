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

## Rust equivalents

| Java POI test | Rust equivalent |
|---------------|-----------------|
| Ehcache probes | `easyexcel_reader::cache::Ehcache` + `ReadCacheMode::Disk` |
| XLS edge cases | `biff8/template.rs` BIFF record tests |
| XLSX edge cases | `xlsx_rows/tests.rs` SAX parser tests |
| Encryption probes | `biff8/encrypt.rs` RC4 encryption tests |
| Format probes | `builtin_formats.rs` format table tests |

These Rust equivalents test the equivalent engine behavior
without depending on POI internals.