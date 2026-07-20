//! Mirrors Java `com.alibaba.excel.write.style.*`.
//!
//! ## Wiring status (annotation metadata vs registered strategies)
//!
//! Annotation-driven styles on [`easyexcel_core::ExcelWriteMetadata`] /
//! [`easyexcel_core::ExcelColumn`] (via `#[derive(ExcelRow)]`) **are** consumed
//! by the XLSX write path (`apply_annotation_*`, `SheetStyleContext`).
//!
//! Registered strategy handlers are also applied by
//! `write_xlsx_with_handlers` / `append_rows_to_worksheet` via the
//! [`easyexcel_core::WriteHandler`] `style_*` accessors:
//!
//! | Strategy | Type | Applied via `register_write_handler` |
//! |---|---|---|
//! | `HorizontalCellStyleStrategy` | yes | **wired** (`style_cell_style` + nested `WriteFont`/`ExcelFontStyle`) |
//! | `VerticalCellStyleStrategy` | yes | **wired** (`AbstractVerticalCellStyleStrategy` + font) |
//! | `LongestMatchColumnWidthStyleStrategy` | yes | **wired** (byte-length `style_column_width`; optional autofit) |
//! | `SimpleColumnWidthStyleStrategy` | yes | **wired** (`style_column_width`) |
//! | `SimpleRowHeightStyleStrategy` | yes | **wired** (`style_*_row_height`) |
//! | `LoopMergeStrategy` | yes | **wired** (`WriteOptions::loop_merges` + `@ContentLoopMerge`) |
//! | `OnceAbsoluteMergeStrategy` | yes | **wired** (`style_once_absolute_merge` + `@OnceAbsoluteMerge`) |

pub mod abstract_cell_style_strategy;
pub mod abstract_vertical_cell_style_strategy;
pub mod default_style;
pub mod horizontal_cell_style_strategy;
pub mod vertical_cell_style_strategy;

pub mod column;
pub mod row;
