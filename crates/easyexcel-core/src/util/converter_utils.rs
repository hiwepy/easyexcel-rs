//! Mirrors Java com.alibaba.excel.util.ConverterUtils.

#![allow(dead_code)]

use std::any::TypeId;

use crate::excel_error::ExcelError;

/// Mirrors `com.alibaba.excel.util.ConverterUtils#convertToJavaObject`.
///
/// The Rust port performs cell-to-field conversion via the
/// `FromExcelCell` trait; this function is the Java-API-shaped anchor
/// returning an `Unsupported` error until wired in by the reader crate.
pub fn convert_to_java_object(_source: &str, _target_type: TypeId) -> Result<String, ExcelError> {
    Err(ExcelError::Unsupported(
        "ConverterUtils.convertToJavaObject: use the FromExcelCell trait instead".to_owned(),
    ))
}

/// Mirrors `com.alibaba.excel.util.ConverterUtils#convertToStringMap`.
///
/// Converts a flat `(key, value)` iterator into a `HashMap<String, String>`,
/// the Rust analogue of the Java `Map<String, String>` produced by the
/// original helper.
#[must_use]
pub fn convert_to_string_map<'a, K, V, I>(entries: I) -> std::collections::HashMap<String, String>
where
    K: AsRef<str> + 'a,
    V: ToString + 'a,
    I: IntoIterator<Item = (&'a K, &'a V)>,
{
    entries
        .into_iter()
        .map(|(k, v)| (k.as_ref().to_owned(), v.to_string()))
        .collect()
}

/// Mirrors `com.alibaba.excel.util.ConverterUtils#defaultClassGeneric`.
#[must_use]
pub fn default_class_generic(_type_id: TypeId) -> Option<TypeId> {
    None
}
