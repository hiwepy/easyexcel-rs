//! Mirrors Java `com.alibaba.excel.converters.NullableObjectConverter`.

/// Mirrors Java `NullableObjectConverter<T> extends Converter<T>`.
///
/// Java uses this marker interface to tell the executor that
/// `convertToExcelData` can accept `null` values. Rust handles `None`
/// through `impl IntoExcelCell for Option<T>`.
#[allow(dead_code)]
pub struct NullableObjectConverter;
