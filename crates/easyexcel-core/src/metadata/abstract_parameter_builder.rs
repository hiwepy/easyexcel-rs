//! Mirrors Java `com.alibaba.excel.metadata.AbstractParameterBuilder`.

use crate::CacheLocation;

use super::basic_parameter::BasicParameter;

/// Shared fluent builder surface for read/write parameter objects.
///
/// Java uses `AbstractParameterBuilder<T, C extends BasicParameter>`. Rust
/// exposes the same method names through a trait so writer/reader builders can
/// reuse the parameter bag without duplicating setter logic.
///
/// Rust port of Java `AbstractParameterBuilder`.
pub trait AbstractParameterBuilder {
    /// Returns the parameter being mutated. (Java `parameter()`)
    fn parameter(&mut self) -> &mut BasicParameter;

    /// Sets dynamic header rows. (Java `head(List<List<String>>)`)
    fn head(&mut self, head: Vec<Vec<String>>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().head = Some(head);
        self
    }

    /// Sets the model type name. (Java `head(Class<?>)` → `setClazz`)
    fn head_class(&mut self, clazz: impl Into<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().clazz = Some(clazz.into());
        self
    }

    /// Registers a custom converter by type name. (Java `registerConverter`)
    fn register_converter(&mut self, converter_type: impl Into<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter()
            .custom_converter_list
            .push(converter_type.into());
        self
    }

    /// Sets the 1904 date-windowing flag. (Java `use1904windowing(Boolean)`)
    fn use1904windowing(&mut self, enabled: bool) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().use1904windowing = Some(enabled);
        self
    }

    /// Sets the formatting locale name. (Java `locale(Locale)`)
    fn locale(&mut self, locale: impl Into<String>) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().locale = Some(locale.into());
        self
    }

    /// Sets the field-cache location. (Java `filedCacheLocation`)
    fn filed_cache_location(&mut self, location: CacheLocation) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().filed_cache_location = Some(location);
        self
    }

    /// Sets the auto-trim flag. (Java `autoTrim(Boolean)`)
    fn auto_trim(&mut self, auto_trim: bool) -> &mut Self
    where
        Self: Sized,
    {
        self.parameter().auto_trim = Some(auto_trim);
        self
    }
}

/// Minimal builder used by metadata tests and future reader/writer facades.
#[derive(Debug, Clone, Default)]
pub struct BasicParameterBuilder {
    parameter: BasicParameter,
}

impl BasicParameterBuilder {
    /// Creates an empty builder. (Java builder entry point)
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Builds the parameter bag. (Java `build()` parameter extraction)
    #[must_use]
    pub fn build(self) -> BasicParameter {
        self.parameter
    }
}

impl AbstractParameterBuilder for BasicParameterBuilder {
    fn parameter(&mut self) -> &mut BasicParameter {
        &mut self.parameter
    }
}
