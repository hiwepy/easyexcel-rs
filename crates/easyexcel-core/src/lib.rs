//! Core data model and extension points for `easyexcel-rs`.
//!
//! This crate is the Rust port of Alibaba `EasyExcel` 4.0.3's
//! `com.alibaba.excel.*` package. Java splits every concept into its own
//! file; the Rust port mirrors that layout 1:1 by giving every public
//! type its own module below. This file is the public re-export
//! surface, so every consumer (`easyexcel-reader`, `easyexcel-writer`,
//! `easyexcel-template`, `easyexcel`, and the `easyexcel-derive` proc-macro
//! crate) keeps its existing `easyexcel_core::Type` import path.

#![warn(missing_docs)]
#![warn(clippy::pedantic)]
#![deny(unsafe_code)]

/// Arbitrary-precision decimal type used for Java `BigDecimal`-compatible cells.
pub use bigdecimal::BigDecimal;
/// Arbitrary-precision integer type used for Java `BigInteger`-compatible fields.
pub use num_bigint::BigInt;
/// Parsed URL type accepted by the default Java-compatible URL image converter.
pub use url::Url;

// ---------------------------------------------------------------------------
// Enums (Java `com.alibaba.excel.enums.*`)
// ---------------------------------------------------------------------------
mod enum_boolean;
mod enum_byte_order_mark;
mod enum_cache_location;
mod enum_cell_data_type;
mod enum_cell_extra_type;
mod enum_head_kind;
mod enum_holder;
mod enum_numeric_cell_type;
mod enum_read_default_return;
mod enum_row_type;
mod enum_write_direction;
mod enum_write_last_row;
mod enum_write_template_analysis_cell_type;
mod enum_write_type;

pub use enum_boolean::*;
pub use enum_byte_order_mark::*;
pub use enum_cache_location::*;
pub use enum_cell_data_type::*;
pub use enum_cell_extra_type::*;
pub use enum_head_kind::*;
pub use enum_holder::*;
pub use enum_numeric_cell_type::*;
pub use enum_read_default_return::*;
pub use enum_row_type::*;
pub use enum_write_direction::*;
pub use enum_write_last_row::*;
pub use enum_write_template_analysis_cell_type::*;
pub use enum_write_type::*;

// ---------------------------------------------------------------------------
// Cell data model (Java `com.alibaba.excel.metadata.data.*` +
//                 `com.alibaba.excel.metadata.*`)
// ---------------------------------------------------------------------------
mod anchor_type;
mod cell_extra;
mod cell_value;
mod client_anchor_data;
mod comment_data;
mod coordinate_data;
mod formula_data;
mod hyperlink_data;
mod image_data;
mod image_type;
mod interval_font;
mod read_cell_data;
mod rich_text_string_data;
mod write_cell_data;
mod write_font;

pub use anchor_type::*;
pub use cell_extra::*;
pub use cell_value::*;
pub use client_anchor_data::*;
pub use comment_data::*;
pub use coordinate_data::*;
pub use formula_data::*;
pub use hyperlink_data::*;
pub use image_data::*;
pub use image_type::*;
pub use interval_font::*;
pub use read_cell_data::*;
pub use rich_text_string_data::*;
pub use write_cell_data::*;
pub use write_font::*;

// ---------------------------------------------------------------------------
// Style metadata (Java `com.alibaba.excel.write.metadata.style.*` + `enums.poi.*`)
// ---------------------------------------------------------------------------
mod excel_border_style;
mod excel_cell_style;
mod excel_color;
mod excel_column;
mod excel_data_format;
mod excel_fill_pattern;
mod excel_font_script;
mod excel_font_style;
mod excel_horizontal_alignment;
mod excel_underline;
mod excel_vertical_alignment;
mod excel_write_metadata;

pub use excel_border_style::*;
pub use excel_cell_style::*;
pub use excel_color::*;
pub use excel_column::*;
pub use excel_data_format::*;
pub use excel_fill_pattern::*;
pub use excel_font_script::*;
pub use excel_font_style::*;
pub use excel_horizontal_alignment::*;
pub use excel_underline::*;
pub use excel_vertical_alignment::*;
pub use excel_write_metadata::*;
pub use metadata::property::{
    ExcelDataValidationMeta, LoopMergeProperty, OnceAbsoluteMergeProperty,
};

// ---------------------------------------------------------------------------
// Conversion context + dynamic rows (Java
//   `com.alibaba.excel.metadata.property.ExcelContentProperty` +
//   `com.alibaba.excel.read.listener.ReadListener` payloads)
// ---------------------------------------------------------------------------
mod analysis_context;
mod convert_context;
mod custom_read_object;
mod dynamic_row;
mod dynamic_value;
mod page_read_listener;
mod row_data;

pub use analysis_context::*;
pub use convert_context::*;
pub use custom_read_object::*;
pub use dynamic_row::*;
pub use dynamic_value::*;
pub use page_read_listener::*;
pub use row_data::*;

// ---------------------------------------------------------------------------
// Conversion traits + registry (Java
//   `com.alibaba.excel.converters.Converter` +
//   `com.alibaba.excel.converters.ConverterKeyBuild` +
//   `com.alibaba.excel.converters.DefaultConverterLoader`)
// ---------------------------------------------------------------------------
pub mod converter;
mod converter_registry;
mod from_excel_cell;
mod from_into_impls;
mod into_excel_cell;
mod read_converter_context;
mod write_converter_context;

pub use converter::converter_trait::*;
pub use converter_registry::*;
pub use from_excel_cell::*;
pub use into_excel_cell::*;
// `from_into_impls` is intentionally not re-exported; it only provides
// the trait impls for the built-in scalar/sequence/option types above.
pub use read_converter_context::*;
pub use write_converter_context::*;

// ---------------------------------------------------------------------------
// Built-in image converters (Java
//   `com.alibaba.excel.converters.string.StringImageConverter` +
//   `com.alibaba.excel.converters.inputstream.InputStreamImageConverter` +
//   `com.alibaba.excel.converters.url.UrlImageConverter`)
// ---------------------------------------------------------------------------
mod image_input_stream;
mod input_stream_image_converter;
mod string_image_converter;
mod url_image_converter;

pub use image_input_stream::*;
pub use input_stream_image_converter::*;
pub use string_image_converter::*;
pub use url_image_converter::*;

// ---------------------------------------------------------------------------
// CSV charset (Java `com.alibaba.excel.support.ExcelTypeEnum` + charset glue)
// ---------------------------------------------------------------------------
mod csv_charset;

pub use csv_charset::*;

// ---------------------------------------------------------------------------
// Errors (Java `com.alibaba.excel.exception.*`)
// ---------------------------------------------------------------------------
mod excel_error;

pub use excel_error::*;

/// The result type used by all easyexcel crates.
pub type Result<T> = std::result::Result<T, ExcelError>;

// ---------------------------------------------------------------------------
// Event listener + write handler traits (Java
//   `com.alibaba.excel.read.listener.ReadListener` +
//   `com.alibaba.excel.write.handler.WriteHandler`)
// ---------------------------------------------------------------------------
mod read_listener;
mod write_cell_context;
mod write_context;
mod write_fill_executor;
mod write_handler;
mod write_row_context;

// ---------------------------------------------------------------------------
// Event package (Java `com.alibaba.excel.event.*`)
// ---------------------------------------------------------------------------
pub mod event;
mod write_sheet_context;
mod write_workbook_context;

pub use read_listener::*;
pub use write_cell_context::*;
pub use write_context::*;
pub use write_fill_executor::*;
pub use write_handler::*;
pub use write_row_context::*;
pub use write_sheet_context::*;
pub use write_workbook_context::*;

// ---------------------------------------------------------------------------
// `ExcelRow` derive trait (Java
//   `com.alibaba.excel.metadata.property.ExcelHeadProperty` runtime + the
//   `ModelBuildEventListener`)
// ---------------------------------------------------------------------------
mod excel_row;

pub use excel_row::*;

// ---------------------------------------------------------------------------
// POI enum re-exports (Java `com.alibaba.excel.enums.poi.*`)
// ---------------------------------------------------------------------------
mod enums;

pub use enums::poi;

// ---------------------------------------------------------------------------
// Exception type aliases (Java `com.alibaba.excel.exception.*`)
// ---------------------------------------------------------------------------
pub mod exception;

// ---------------------------------------------------------------------------
// Support (Java `com.alibaba.excel.support.*`)
// ---------------------------------------------------------------------------
pub mod support;

// ---------------------------------------------------------------------------
// Constants (Java `com.alibaba.excel.constant.*`)
// ---------------------------------------------------------------------------
pub mod constant;

// ---------------------------------------------------------------------------
// Metadata sub-packages (Java `com.alibaba.excel.metadata.property.*`,
//   `com.alibaba.excel.metadata.format.*`, `com.alibaba.excel.metadata.csv.*`)
// ---------------------------------------------------------------------------
pub mod metadata;

// ---------------------------------------------------------------------------
// Annotations (Java `com.alibaba.excel.annotation.*`)
// ---------------------------------------------------------------------------
pub mod annotation;

// ---------------------------------------------------------------------------
// Utilities (Java `com.alibaba.excel.util.*`)
// ---------------------------------------------------------------------------
pub mod util;

#[cfg(test)]
mod tests;
#[cfg(test)]
mod missing_tests;
