//! Missing test coverage for read-side features.

use super::*;

// ============================================================================
// CacheDataTest (4 tests) — Cache mode
// ============================================================================

#[test]
fn cache_mode_default_is_auto() {
    assert_eq!(ReadCacheMode::default(), ReadCacheMode::Auto);
}

#[test]
fn cache_mode_memory_variant() {
    let mode = ReadCacheMode::Memory;
    assert_eq!(mode, ReadCacheMode::Memory);
    assert_ne!(mode, ReadCacheMode::Disk);
    assert_ne!(mode, ReadCacheMode::Auto);
}

#[test]
fn cache_mode_disk_variant() {
    let mode = ReadCacheMode::Disk;
    assert_eq!(mode, ReadCacheMode::Disk);
    assert_ne!(mode, ReadCacheMode::Memory);
    assert_ne!(mode, ReadCacheMode::Auto);
}

#[test]
fn cache_mode_all_variants_distinct() {
    assert_ne!(ReadCacheMode::Auto, ReadCacheMode::Memory);
    assert_ne!(ReadCacheMode::Auto, ReadCacheMode::Disk);
    assert_ne!(ReadCacheMode::Memory, ReadCacheMode::Disk);
}

// ============================================================================
// MultipleSheetsDataTest (5 tests) — Multiple sheets
// ============================================================================

#[test]
fn multiple_sheets_indices() {
    let s0 = SheetSelector::Index(0);
    let s1 = SheetSelector::Index(1);
    let s2 = SheetSelector::Index(2);
    assert_eq!(s0, SheetSelector::Index(0));
    assert_eq!(s1, SheetSelector::Index(1));
    assert_eq!(s2, SheetSelector::Index(2));
    assert_ne!(s0, s1);
    assert_ne!(s1, s2);
}

#[test]
fn multiple_sheets_names() {
    let s1 = SheetSelector::Name("Sheet1".to_owned());
    let s2 = SheetSelector::Name("Data".to_owned());
    assert_eq!(s1, SheetSelector::Name("Sheet1".to_owned()));
    assert_eq!(s2, SheetSelector::Name("Data".to_owned()));
    assert_ne!(s1, s2);
}

#[test]
fn multiple_sheets_default_is_first() {
    assert_eq!(SheetSelector::default(), SheetSelector::First);
}

#[test]
fn multiple_sheets_all_selector() {
    let all = SheetSelector::All;
    assert_eq!(all, SheetSelector::All);
}

#[test]
fn multiple_sheets_distinct_types() {
    assert_ne!(
        SheetSelector::Index(0),
        SheetSelector::Name("Test".to_owned())
    );
    assert_ne!(SheetSelector::All, SheetSelector::First);
}

// ============================================================================
// NoHeadDataTest (4 tests) — No head
// ============================================================================

#[test]
fn no_head_basic_options() {
    let mut opts = ReadOptions::default();
    opts.ignore_empty_row = true;
    opts.auto_trim = false;
    assert!(opts.ignore_empty_row);
    assert!(!opts.auto_trim);
}

#[test]
fn no_head_with_use_1904_windowing() {
    let mut opts = ReadOptions::default();
    opts.use_1904_windowing = true;
    assert!(opts.use_1904_windowing);
}

#[test]
fn no_head_with_scientific_format() {
    let mut opts = ReadOptions::default();
    opts.scientific_format = ScientificFormatMode::Scientific;
    assert_eq!(opts.scientific_format, ScientificFormatMode::Scientific);
}

#[test]
fn no_head_default_options() {
    let opts = ReadOptions::default();
    assert!(opts.ignore_empty_row);
    assert!(opts.auto_trim);
    assert!(!opts.use_1904_windowing);
}

// ============================================================================
// ListHeadDataTest (4 tests) — List head
// ============================================================================

#[test]
fn list_head_dynamic_default() {
    let mut opts = ReadOptions::default();
    opts.head_row_number = 2;
    assert_eq!(opts.head_row_number, 2);
}

#[test]
fn list_head_with_single_level() {
    let opts = ReadOptions::default();
    assert_eq!(opts.head_row_number, 1);
}

#[test]
fn list_head_with_three_levels() {
    let mut opts = ReadOptions::default();
    opts.head_row_number = 3;
    assert_eq!(opts.head_row_number, 3);
}

#[test]
fn list_head_with_zero() {
    let mut opts = ReadOptions::default();
    opts.head_row_number = 0;
    assert_eq!(opts.head_row_number, 0);
}

// ============================================================================
// CellDataDataTest (4 tests) — Cell data type (read side)
// ============================================================================

#[test]
fn cell_data_type_string_default() {
    let opts = ReadOptions::default();
    assert_eq!(opts.read_default_return, ReadDefaultReturn::String);
}

#[test]
fn cell_data_type_actual_data_mode() {
    let mut opts = ReadOptions::default();
    opts.read_default_return = ReadDefaultReturn::ActualData;
    assert_eq!(opts.read_default_return, ReadDefaultReturn::ActualData);
}

#[test]
fn cell_data_type_read_cell_data_mode() {
    let mut opts = ReadOptions::default();
    opts.read_default_return = ReadDefaultReturn::ReadCellData;
    assert_eq!(opts.read_default_return, ReadDefaultReturn::ReadCellData);
}

#[test]
fn cell_data_type_all_modes() {
    assert_ne!(ReadDefaultReturn::String, ReadDefaultReturn::ActualData);
    assert_ne!(ReadDefaultReturn::String, ReadDefaultReturn::ReadCellData);
    assert_ne!(
        ReadDefaultReturn::ActualData,
        ReadDefaultReturn::ReadCellData
    );
}
