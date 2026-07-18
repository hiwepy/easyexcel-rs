//! Mirrors Java `com.alibaba.excel.converters.ConverterKeyBuild`.

use crate::enum_cell_data_type::CellDataType;

/// Builds a converter dispatch key from a Rust type name and Excel cell type.
/// (Java `ConverterKeyBuild.buildKey(Class, CellDataTypeEnum)`)
#[allow(dead_code)]
pub fn build_key(type_name: &str, cell_data_type: Option<CellDataType>) -> String {
    match cell_data_type {
        Some(t) => format!("{type_name}:{t:?}"),
        None => type_name.to_owned(),
    }
}
