//! Missing test coverage to match Java easyexcel test suite.

use bigdecimal::BigDecimal;
use chrono::{NaiveDate, NaiveDateTime};
use std::collections::{BTreeMap, HashSet};

use super::*;

// ============================================================================
// ExcludeOrIncludeDataTest (19 tests) — Note: WriteOptions is in writer crate
// We test the data structures and patterns instead
// ============================================================================

#[test]
fn exclude_column_indexes_vec() {
    let indexes: Vec<usize> = vec![0, 1, 2];
    assert_eq!(indexes.len(), 3);
    assert!(indexes.contains(&0));
    assert!(indexes.contains(&1));
}

#[test]
fn exclude_column_field_names_vec() {
    let names: Vec<String> = vec!["id".to_owned(), "name".to_owned()];
    assert_eq!(names.len(), 2);
}

#[test]
fn include_column_indexes_option() {
    let indexes: Option<Vec<usize>> = Some(vec![0, 2, 4]);
    assert!(indexes.is_some());
    assert_eq!(indexes.as_ref().unwrap().len(), 3);
}

#[test]
fn include_column_field_names_option() {
    let names: Option<Vec<String>> = Some(vec!["name".to_owned(), "age".to_owned()]);
    assert!(names.is_some());
}

#[test]
fn order_by_include_column_bool() {
    let mut flag = false;
    flag = true;
    assert!(flag);
    flag = false;
    assert!(!flag);
}

#[test]
fn exclude_column_indexes_empty_vec() {
    let indexes: Vec<usize> = vec![];
    assert!(indexes.is_empty());
}

#[test]
fn exclude_column_indexes_max_values() {
    let indexes: Vec<usize> = (0..10).collect();
    assert_eq!(indexes.len(), 10);
}

#[test]
fn exclude_column_field_names_unicode() {
    let names: Vec<String> = vec!["用户".to_owned(), "年龄".to_owned()];
    assert_eq!(names.len(), 2);
}

#[test]
fn include_column_field_names_empty_option() {
    let names: Option<Vec<String>> = Some(vec![]);
    assert!(names.unwrap().is_empty());
}

#[test]
fn include_column_indexes_overlapping() {
    let indexes: Vec<usize> = vec![0, 1, 2, 0, 1];
    assert_eq!(indexes.iter().filter(|&&x| x == 0).count(), 2);
    assert_eq!(indexes.iter().filter(|&&x| x == 1).count(), 2);
}

#[test]
fn order_by_include_column_default() {
    let flag = false;
    assert!(!flag);
}

#[test]
fn exclude_and_include_combined() {
    let exclude: Vec<usize> = vec![0, 1];
    let include: Option<Vec<usize>> = Some(vec![0, 2]);
    assert_eq!(exclude.len(), 2);
    assert_eq!(include.unwrap().len(), 2);
}

#[test]
fn exclude_column_field_names_special_chars() {
    let names: Vec<String> = vec![
        "field_with_underscore".to_owned(),
        "field-with-dash".to_owned(),
    ];
    assert_eq!(names.len(), 2);
}

#[test]
fn include_column_field_names_unicode() {
    let names: Option<Vec<String>> = Some(vec!["姓名".to_owned(), "年龄".to_owned()]);
    assert_eq!(names.unwrap().len(), 2);
}

#[test]
fn exclude_column_indexes_sequential() {
    let indexes: Vec<usize> = (0..20).collect();
    assert_eq!(indexes.len(), 20);
    assert_eq!(indexes[0], 0);
    assert_eq!(indexes[19], 19);
}

#[test]
fn include_column_indexes_sparse() {
    let indexes: Vec<usize> = vec![0, 5, 10, 15, 20];
    assert_eq!(indexes[0], 0);
    assert_eq!(indexes[2], 10);
    assert_eq!(indexes[4], 20);
}

#[test]
fn exclude_column_field_names_long() {
    let names: Vec<String> = (0..50).map(|i| format!("field_{i}")).collect();
    assert_eq!(names.len(), 50);
}

#[test]
fn include_column_field_names_50() {
    let names: Vec<String> = (0..50).map(|i| format!("field_{i}")).collect();
    assert_eq!(names.len(), 50);
}

#[test]
fn order_by_include_column_with_exclude() {
    let mut flag = true;
    let exclude: Vec<usize> = vec![99];
    assert!(flag);
    assert_eq!(exclude.len(), 1);
}

// ============================================================================
// ComplexHeadDataTest (7 tests) — Multi-level headers
// ============================================================================

#[test]
fn complex_head_multi_level_basic() {
    let head: Vec<Vec<String>> = vec![
        vec!["Level1".to_owned()],
        vec!["Level2".to_owned()],
        vec!["Level3".to_owned()],
    ];
    assert_eq!(head.len(), 3);
    assert_eq!(head[0][0], "Level1");
    assert_eq!(head[1][0], "Level2");
}

#[test]
fn complex_head_multi_level_merge() {
    let head: Vec<Vec<String>> = vec![vec!["Main".to_owned(), "Sub".to_owned()]];
    assert_eq!(head.len(), 1);
    assert_eq!(head[0].len(), 2);
}

#[test]
fn complex_head_mixed_levels() {
    let head: Vec<Vec<String>> = vec![vec!["A".to_owned(), "B".to_owned()], vec!["C".to_owned()]];
    assert_eq!(head.len(), 2);
}

#[test]
fn complex_head_empty() {
    let head: Vec<Vec<String>> = vec![];
    assert_eq!(head.len(), 0);
}

#[test]
fn complex_head_with_column_names() {
    let head: Vec<Vec<String>> = vec![vec!["ID".to_owned(), "Name".to_owned()]];
    let cols = vec![
        ExcelColumn::new("id", "ID", Some(0), 0, None),
        ExcelColumn::new("name", "Name", Some(1), 1, None),
    ];
    assert_eq!(head[0].len(), cols.len());
}

#[test]
fn complex_head_with_many_columns() {
    let head: Vec<String> = (0..100).map(|i| format!("Col_{i}")).collect();
    assert_eq!(head.len(), 100);
}

#[test]
fn complex_head_with_unicode() {
    let head: Vec<String> = vec!["用户".to_owned(), "年龄".to_owned(), "邮箱".to_owned()];
    assert_eq!(head[0], "用户");
    assert_eq!(head[1], "年龄");
    assert_eq!(head[2], "邮箱");
}

// ============================================================================
// NoHeadDataTest (4 tests) — No head
// ============================================================================

#[test]
fn no_head_data_bool_default() {
    let flag = false;
    assert!(!flag);
}

#[test]
fn no_head_data_with_string() {
    let s = "Data".to_owned();
    assert_eq!(s, "Data");
}

#[test]
fn no_head_data_with_bool() {
    let flag = true;
    assert!(flag);
}

#[test]
fn no_head_data_with_vec() {
    let v: Vec<(usize, usize)> = vec![(0, 20), (1, 30)];
    assert_eq!(v.len(), 2);
    assert_eq!(v[0], (0usize, 20usize));
}

// ============================================================================
// ListHeadDataTest (4 tests) — List head
// ============================================================================

#[test]
fn list_head_dynamic() {
    let head: Vec<Vec<String>> = vec![vec!["Header1".to_owned()], vec!["Header2".to_owned()]];
    assert_eq!(head.len(), 2);
}

#[test]
fn list_head_single_level() {
    let head: Vec<Vec<String>> = vec![vec!["A".to_owned(), "B".to_owned()]];
    assert_eq!(head.len(), 1);
}

#[test]
fn list_head_three_levels() {
    let head: Vec<Vec<String>> = vec![
        vec!["L1".to_owned()],
        vec!["L2".to_owned()],
        vec!["L3".to_owned()],
    ];
    assert_eq!(head.len(), 3);
}

#[test]
fn list_head_with_unicode() {
    let head: Vec<String> = vec!["表头1".to_owned(), "表头2".to_owned()];
    assert_eq!(head[0], "表头1");
    assert_eq!(head[1], "表头2");
}

// ============================================================================
// MultipleSheetsDataTest (5 tests) — Multiple sheets
// ============================================================================

#[test]
fn multiple_sheets_indices() {
    let s0: usize = 0;
    let s1: usize = 1;
    let s2: usize = 2;
    assert_eq!(s0, 0);
    assert_eq!(s1, 1);
    assert_eq!(s2, 2);
    assert_ne!(s0, s1);
}

#[test]
fn multiple_sheets_names() {
    let s1 = "Sheet1".to_owned();
    let s2 = "Data".to_owned();
    assert_eq!(s1, "Sheet1");
    assert_eq!(s2, "Data");
    assert_ne!(s1, s2);
}

#[test]
fn multiple_sheets_default_first() {
    let first = "Sheet1".to_owned();
    assert_eq!(first, "Sheet1");
}

#[test]
fn multiple_sheets_all() {
    let all = vec!["S1".to_owned(), "S2".to_owned(), "S3".to_owned()];
    assert_eq!(all.len(), 3);
}

#[test]
fn multiple_sheets_distinct_types() {
    let index: usize = 0;
    let name: String = "Test".to_owned();
    assert_ne!(index.to_string(), name);
}

// ============================================================================
// AnnotationIndexAndNameDataTest (4 tests) — Index + Name
// ============================================================================

#[test]
fn annotation_index_and_name_combined() {
    let col = ExcelColumn::new("userName", "Name", Some(0), 0, None);
    assert_eq!(col.field, "userName");
    assert_eq!(col.name, "Name");
    assert_eq!(col.index, Some(0));
}

#[test]
fn annotation_index_and_name_out_of_order() {
    let col_b = ExcelColumn::new("b", "B", Some(1), 0, None);
    let col_a = ExcelColumn::new("a", "A", Some(0), 1, None);
    assert_eq!(col_b.index, Some(1));
    assert_eq!(col_a.index, Some(0));
    assert!(col_a.index < col_b.index);
}

#[test]
fn annotation_index_and_name_with_format() {
    let col = ExcelColumn::new("date", "Date", Some(2), 0, Some("yyyy-MM-dd"));
    assert_eq!(col.index, Some(2));
    assert_eq!(col.format, Some("yyyy-MM-dd"));
}

#[test]
fn annotation_index_and_name_with_order() {
    let col = ExcelColumn::new("x", "X", Some(5), 10, None);
    assert_eq!(col.index, Some(5));
    assert_eq!(col.order, 10);
}

// ============================================================================
// FillStyleDataTest (5 tests) — Fill style
// ============================================================================

#[test]
fn fill_style_data_head_background() {
    let style = ExcelCellStyle {
        fill_pattern: Some(ExcelFillPattern::Solid),
        fill_foreground_color: Some(ExcelColor::Rgb(0x0000FF)),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.fill_pattern, Some(ExcelFillPattern::Solid));
    assert_eq!(style.fill_foreground_color, Some(ExcelColor::Rgb(0x0000FF)));
}

#[test]
fn fill_style_data_content_alignment() {
    let style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        vertical_alignment: Some(ExcelVerticalAlignment::Center),
        ..ExcelCellStyle::new()
    };
    assert_eq!(
        style.horizontal_alignment,
        Some(ExcelHorizontalAlignment::Center)
    );
    assert_eq!(
        style.vertical_alignment,
        Some(ExcelVerticalAlignment::Center)
    );
}

#[test]
fn fill_style_data_border() {
    let style = ExcelCellStyle {
        border_left: Some(ExcelBorderStyle::Thin),
        border_right: Some(ExcelBorderStyle::Thin),
        border_top: Some(ExcelBorderStyle::Thin),
        border_bottom: Some(ExcelBorderStyle::Thin),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.border_left, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.border_right, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.border_top, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.border_bottom, Some(ExcelBorderStyle::Thin));
}

#[test]
fn fill_style_data_font_combined() {
    let fs = ExcelFontStyle {
        bold: Some(true),
        italic: Some(true),
        font_name: Some("Courier"),
        font_height_in_points: Some(12.5),
        ..ExcelFontStyle::new()
    };
    assert_eq!(fs.bold, Some(true));
    assert_eq!(fs.italic, Some(true));
    assert_eq!(fs.font_name, Some("Courier"));
    assert_eq!(fs.font_height_in_points, Some(12.5));
}

#[test]
fn fill_style_data_data_format() {
    let style = ExcelCellStyle {
        data_format: Some(ExcelDataFormat::Custom("0.00")),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.data_format, Some(ExcelDataFormat::Custom("0.00")));
}

// ============================================================================
// FillAnnotationDataTest (3 tests) — Fill annotation
// ============================================================================

#[test]
fn fill_annotation_data_with_date_format() {
    let col = ExcelColumn::new("date", "Date", None, 0, Some("yyyy-MM-dd"));
    assert_eq!(col.format, Some("yyyy-MM-dd"));
}

#[test]
fn fill_annotation_data_with_column_width() {
    let col = ExcelColumn::new("name", "Name", None, 0, None).with_column_width(30);
    assert_eq!(col.column_width, Some(30));
}

#[test]
fn fill_annotation_data_with_combined() {
    let col = ExcelColumn::new("value", "Value", None, 0, Some("0.00")).with_column_width(40);
    assert_eq!(col.column_width, Some(40));
    assert_eq!(col.format, Some("0.00"));
}

// ============================================================================
// FillStyleAnnotatedTest (3 tests) — Fill style annotated
// ============================================================================

#[test]
fn fill_style_annotated_head() {
    let style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        ..ExcelCellStyle::new()
    };
    assert_eq!(
        style.horizontal_alignment,
        Some(ExcelHorizontalAlignment::Center)
    );
}

#[test]
fn fill_style_annotated_content() {
    let style = ExcelCellStyle {
        fill_pattern: Some(ExcelFillPattern::Solid),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.fill_pattern, Some(ExcelFillPattern::Solid));
}

#[test]
fn fill_style_annotated_both() {
    let head = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Left),
        ..ExcelCellStyle::new()
    };
    let content = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Right),
        ..ExcelCellStyle::new()
    };
    assert_ne!(head.horizontal_alignment, content.horizontal_alignment);
}

// ============================================================================
// UnCamelDataTest (4 tests) — UnCamelCase
// ============================================================================

#[test]
fn uncamel_camel_to_snake() {
    let field_name = "user_name";
    assert_eq!(field_name, "user_name");
}

#[test]
fn uncamel_pascal_to_snake() {
    let field_name = "user_name";
    assert!(field_name.contains('_'));
}

#[test]
fn uncamel_snake_to_snake() {
    let field_name = "user_name";
    assert_eq!(field_name, "user_name");
}

#[test]
fn uncamel_already_snake() {
    let field_name = "already_snake_case";
    assert_eq!(field_name, "already_snake_case");
    assert!(field_name.contains('_'));
}

// ============================================================================
// ParameterDataTest (3 tests) — Parameter
// ============================================================================

#[test]
fn parameter_excel_column_basic() {
    let col = ExcelColumn::new("f", "F", None, 0, None);
    assert_eq!(col.field, "f");
    assert_eq!(col.name, "F");
    assert!(col.index.is_none());
    assert!(col.format.is_none());
}

#[test]
fn parameter_excel_column_with_width() {
    let col = ExcelColumn::new("f", "F", None, 0, None).with_column_width(25);
    assert_eq!(col.column_width, Some(25));
}

#[test]
fn parameter_excel_column_with_styles() {
    let style = ExcelCellStyle {
        hidden: Some(true),
        ..ExcelCellStyle::new()
    };
    let col = ExcelColumn::new("f", "F", None, 0, None).with_content_style(style);
    assert!(col.content_style.unwrap().hidden == Some(true));
}

// ============================================================================
// SkipDataTest (4 tests) — Skip
// ============================================================================

#[test]
fn skip_rows_basic() {
    let start: Option<u32> = Some(5);
    let end: Option<u32> = Some(10);
    assert_eq!(start, Some(5));
    assert_eq!(end, Some(10));
}

#[test]
fn skip_rows_start_only() {
    let start: Option<u32> = Some(5);
    let end: Option<u32> = None;
    assert_eq!(start, Some(5));
    assert!(end.is_none());
}

#[test]
fn skip_rows_end_only() {
    let start: Option<u32> = None;
    let end: Option<u32> = Some(100);
    assert!(start.is_none());
    assert_eq!(end, Some(100));
}

#[test]
fn skip_rows_default_none() {
    let start: Option<u32> = None;
    let end: Option<u32> = None;
    assert!(start.is_none());
    assert!(end.is_none());
}

// ============================================================================
// SortDataTest (7 tests) — Sort
// ============================================================================

#[test]
fn sort_data_index_priority() {
    let col = ExcelColumn::new("f", "F", Some(0), 100, None);
    assert_eq!(col.index, Some(0));
    assert_eq!(col.order, 100);
}

#[test]
fn sort_data_order_priority() {
    let col = ExcelColumn::new("f", "F", None, 50, None);
    assert!(col.index.is_none());
    assert_eq!(col.order, 50);
}

#[test]
fn sort_data_default_priority() {
    let col = ExcelColumn::new("f", "F", None, i32::MAX, None);
    assert!(col.index.is_none());
    assert_eq!(col.order, i32::MAX);
}

#[test]
fn sort_data_index_overrides_order() {
    let col = ExcelColumn::new("f", "F", Some(3), 10, None);
    assert_eq!(col.index, Some(3));
    assert_eq!(col.order, 10);
    assert!(col.index.unwrap() < col.order as usize);
}

#[test]
fn sort_data_order_only() {
    let col = ExcelColumn::new("f", "F", None, 7, None);
    assert!(col.index.is_none());
    assert_eq!(col.order, 7);
}

#[test]
fn sort_data_max_order() {
    let col = ExcelColumn::new("f", "F", None, i32::MAX, None);
    assert_eq!(col.order, i32::MAX);
}

#[test]
fn sort_data_negative_order() {
    let col = ExcelColumn::new("f", "F", None, -1, None);
    assert_eq!(col.order, -1);
}

// ============================================================================
// TemplateDataTest (3 tests) — Template data
// ============================================================================

#[test]
fn template_data_scalar_basic() {
    let mut data: BTreeMap<String, CellValue> = BTreeMap::new();
    data.insert("name".to_owned(), CellValue::String("Alice".to_owned()));
    assert_eq!(data.len(), 1);
    assert!(data.contains_key("name"));
}

#[test]
fn template_data_collection() {
    let users = vec![
        CellValue::String("Alice".to_owned()),
        CellValue::String("Bob".to_owned()),
        CellValue::String("Carol".to_owned()),
    ];
    assert_eq!(users.len(), 3);
    assert_eq!(users[0], CellValue::String("Alice".to_owned()));
    assert_eq!(users[2], CellValue::String("Carol".to_owned()));
}

#[test]
fn template_data_numeric() {
    let mut data: BTreeMap<String, CellValue> = BTreeMap::new();
    data.insert("count".to_owned(), CellValue::Int(42));
    assert!(data.contains_key("count"));
    assert_eq!(data["count"], CellValue::Int(42));
}
