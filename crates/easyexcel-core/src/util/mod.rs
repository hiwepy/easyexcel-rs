//! Mirrors Java `com.alibaba.excel.util.*`.
//!
//! Each submodule is a 1:1 mirror of a single Java utility class from
//! `com.alibaba.excel.util` (and the Apache Commons / POI helpers it
//! delegates to). File names are the `snake_case` form of the Java
//! class name; every public static method becomes a `pub fn`.

pub mod bean_map_utils;
pub mod boolean_utils;
pub mod class_utils;
pub mod converter_utils;
pub mod date_utils;
pub mod easy_excel_temp_file_creation_strategy;
pub mod field_utils;
pub mod file_type_utils;
pub mod file_utils;
pub mod int_utils;
pub mod io_utils;
pub mod list_utils;
pub mod map_utils;
pub mod member_utils;
pub mod number_data_formatter_utils;
pub mod number_utils;
pub mod poi_utils;
pub mod position_utils;
pub mod sheet_utils;
pub mod string_utils;
pub mod style_util;
pub mod validate;
pub mod work_book_util;
pub mod write_handler_utils;

// Note: we intentionally do NOT flatten via `pub use xxx::*;` because
// several submodules re-declare the same Java method name
// (`format`, `remove_thread_local_cache`, ...). Callers should reach
// for a helper through its fully-qualified path, e.g.
// `easyexcel_core::util::string_utils::is_blank`.
