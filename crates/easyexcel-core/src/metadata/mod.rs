//! Mirrors Java `com.alibaba.excel.metadata.*` sub-packages.

pub mod abstract_cell;
pub mod abstract_holder;
pub mod abstract_parameter_builder;
pub mod basic_parameter;
pub mod cell;
pub mod cell_range;
pub mod configuration_holder;
pub mod field_cache;
pub mod field_wrapper;
pub mod fill;
pub mod font;
pub mod holder;
pub mod data;
pub mod global_configuration;
pub mod head;
pub mod null_object;
pub mod property;
pub mod format;
pub mod csv;

#[cfg(test)]
mod tests;

pub use abstract_cell::AbstractCell;
pub use abstract_holder::AbstractHolder;
pub use abstract_parameter_builder::{AbstractParameterBuilder, BasicParameterBuilder};
pub use basic_parameter::BasicParameter;
pub use cell::Cell;
pub use cell_range::CellRange;
pub use configuration_holder::{ConfigurationHolder, MetadataHolder};
pub use field_cache::FieldCache;
pub use field_wrapper::FieldWrapper;
pub use fill::AnalysisCell;
pub use global_configuration::GlobalConfiguration;
pub use head::Head;
pub use null_object::NullObject;

pub use property::{
    ColumnWidthProperty, DateTimeFormatProperty, ExcelContentProperty, ExcelHeadProperty,
    ExcelReadHeadProperty, FontProperty, LoopMergeProperty, NumberFormatProperty,
    OnceAbsoluteMergeProperty, RowHeightProperty, StyleProperty,
};

pub use font::Font;
pub use holder::{ExcelHolder, HolderEnum};
pub use data::{CellData, DataFormatData};
