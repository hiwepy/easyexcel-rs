//! Missing test coverage to match Java easyexcel test suite.
//!
//! Tests that use types from easyexcel-writer (LoopMergeStrategy, etc.)

use super::*;

// ============================================================================
// RepetitionDataTest (7 tests) — Repetition
// ============================================================================

#[test]
fn repetition_loop_merge_basic_each_2() {
    // Java: @ContentLoopMerge(eachRow = 2, columnExtend = 1)
    let strategy = LoopMergeStrategy::new(2, 1, 0).unwrap();
    assert_eq!(strategy.each_rows(), 2);
    assert_eq!(strategy.column_extend(), 1);
    assert_eq!(strategy.column_index(), 0);
}

#[test]
fn repetition_loop_merge_each_3_extend_2() {
    let strategy = LoopMergeStrategy::new(3, 2, 1).unwrap();
    assert_eq!(strategy.each_rows(), 3);
    assert_eq!(strategy.column_extend(), 2);
}

#[test]
fn repetition_loop_merge_zero_index() {
    let strategy = LoopMergeStrategy::new(2, 1, 0).unwrap();
    assert_eq!(strategy.column_index(), 0);
}

#[test]
fn repetition_loop_merge_high_index() {
    let strategy = LoopMergeStrategy::new(2, 1, 99).unwrap();
    assert_eq!(strategy.column_index(), 99);
}

#[test]
fn repetition_loop_merge_max_extend() {
    let strategy = LoopMergeStrategy::new(2, u16::MAX, 0).unwrap();
    assert_eq!(strategy.column_extend(), u16::MAX);
}

#[test]
fn repetition_loop_merge_large_each() {
    let strategy = LoopMergeStrategy::new(1000, 5, 0).unwrap();
    assert_eq!(strategy.each_rows(), 1000);
    assert_eq!(strategy.column_extend(), 5);
}

#[test]
fn repetition_loop_merge_all_fields_distinct() {
    let s1 = LoopMergeStrategy::new(2, 1, 0).unwrap();
    let s2 = LoopMergeStrategy::new(3, 2, 1).unwrap();
    assert_ne!(s1.each_rows(), s2.each_rows());
    assert_ne!(s1.column_extend(), s2.column_extend());
    assert_ne!(s1.column_index(), s2.column_index());
}

// ============================================================================
// FillStyleDataTest (5 tests) — Fill style
// ============================================================================

#[test]
fn fill_style_data_head_background() {
    use easyexcel_core::{ExcelCellStyle, ExcelColor, ExcelFillPattern, ExcelHorizontalAlignment, ExcelVerticalAlignment, ExcelBorderStyle, ExcelFontStyle};
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
    use easyexcel_core::{ExcelCellStyle, ExcelHorizontalAlignment, ExcelVerticalAlignment};
    let style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        vertical_alignment: Some(ExcelVerticalAlignment::Center),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.horizontal_alignment, Some(ExcelHorizontalAlignment::Center));
    assert_eq!(style.vertical_alignment, Some(ExcelVerticalAlignment::Center));
}

#[test]
fn fill_style_data_border() {
    use easyexcel_core::{ExcelCellStyle, ExcelBorderStyle};
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
    use easyexcel_core::{ExcelFontStyle, ExcelCellStyle};
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
    use easyexcel_core::{ExcelCellStyle, ExcelDataFormat};
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
    use easyexcel_core::ExcelColumn;
    let col = ExcelColumn::new("date", "Date", None, 0, Some("yyyy-MM-dd"));
    assert_eq!(col.format, Some("yyyy-MM-dd"));
}

#[test]
fn fill_annotation_data_with_column_width() {
    use easyexcel_core::ExcelColumn;
    let col = ExcelColumn::new("name", "Name", None, 0, None).with_column_width(30);
    assert_eq!(col.column_width, Some(30));
}

#[test]
fn fill_annotation_data_with_combined() {
    use easyexcel_core::ExcelColumn;
    let col = ExcelColumn::new("value", "Value", None, 0, Some("0.00"))
        .with_column_width(40);
    assert_eq!(col.column_width, Some(40));
    assert_eq!(col.format, Some("0.00"));
}

// ============================================================================
// FillStyleAnnotatedTest (3 tests) — Fill style annotated
// ============================================================================

#[test]
fn fill_style_annotated_head() {
    use easyexcel_core::{ExcelCellStyle, ExcelHorizontalAlignment};
    let style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.horizontal_alignment, Some(ExcelHorizontalAlignment::Center));
}

#[test]
fn fill_style_annotated_content() {
    use easyexcel_core::{ExcelCellStyle, ExcelFillPattern};
    let style = ExcelCellStyle {
        fill_pattern: Some(ExcelFillPattern::Solid),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.fill_pattern, Some(ExcelFillPattern::Solid));
}

#[test]
fn fill_style_annotated_both() {
    use easyexcel_core::{ExcelCellStyle, ExcelHorizontalAlignment};
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
