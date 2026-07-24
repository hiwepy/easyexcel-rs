//! Mirrors Java `com.alibaba.excel.converters.NullableObjectConverter`.

use super::converter_trait::Converter;

/// Marker for converters that intentionally receive empty cells or absent values.
///
/// Mirrors Java `NullableObjectConverter<T> extends Converter<T>`. Register
/// marker implementations with
/// [`crate::ConverterRegistry::register_nullable`] so the dispatcher can
/// distinguish them from ordinary converters without unstable specialization.
pub trait NullableObjectConverter<T>: Converter<T> {}
