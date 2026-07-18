//! Mirrors Java `com.alibaba.excel.converters.DefaultConverterLoader`.

use crate::converter_registry::ConverterRegistry;

/// Mirrors Java `DefaultConverterLoader.loadDefaultWriteConverter()`.
#[allow(dead_code)]
pub fn load_default_write_converter() -> ConverterRegistry {
    ConverterRegistry::default()
}

/// Mirrors Java `DefaultConverterLoader.loadDefaultReadConverter()`.
#[allow(dead_code)]
pub fn load_default_read_converter() -> ConverterRegistry {
    ConverterRegistry::default()
}
