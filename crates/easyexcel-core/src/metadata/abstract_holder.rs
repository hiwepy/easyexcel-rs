//! Mirrors Java `com.alibaba.excel.metadata.AbstractHolder`.

use crate::CacheLocation;
use crate::ConverterRegistry;
use crate::Holder as HolderEnum;

use super::basic_parameter::BasicParameter;
use super::configuration_holder::ConfigurationHolder;
use super::global_configuration::GlobalConfiguration;

/// Shared holder state for read and write pipelines.
///
/// Rust port of Java `AbstractHolder implements ConfigurationHolder`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AbstractHolder {
    /// Whether the holder was created in this request. (Java `newInitialization`)
    pub new_initialization: bool,
    /// Dynamic header rows. (Java `head`)
    pub head: Option<Vec<Vec<String>>>,
    /// Model type name. (Java `clazz`)
    pub clazz: Option<String>,
    /// Global configuration. (Java `globalConfiguration`)
    pub global_configuration: GlobalConfiguration,
    /// Registered converters. (Java `converterMap`)
    pub converter_map: ConverterRegistry,
    /// Holder scope. (Java `holderType()` on concrete subclasses)
    pub holder_type: HolderEnum,
}

impl Default for AbstractHolder {
    fn default() -> Self {
        Self::new(HolderEnum::Workbook)
    }
}

impl AbstractHolder {
    /// Creates an empty workbook-scoped holder. (Java no-args constructor)
    #[must_use]
    pub fn new(holder_type: HolderEnum) -> Self {
        Self {
            new_initialization: true,
            head: None,
            clazz: None,
            global_configuration: GlobalConfiguration::new(),
            converter_map: ConverterRegistry::default(),
            holder_type,
        }
    }

    /// Initializes holder state from builder parameters and an optional parent.
    /// (Java `AbstractHolder(BasicParameter, AbstractHolder)`)
    #[must_use]
    pub fn from_parameter(
        basic_parameter: &BasicParameter,
        parent: Option<&AbstractHolder>,
        holder_type: HolderEnum,
    ) -> Self {
        let mut holder = Self::new(holder_type);
        holder.new_initialization = true;

        if basic_parameter.head.is_none()
            && basic_parameter.clazz.is_none()
            && let Some(parent) = parent
        {
            holder.head = parent.head.clone();
        } else {
            holder.head = basic_parameter.head.clone();
        }

        if basic_parameter.head.is_none()
            && basic_parameter.clazz.is_none()
            && let Some(parent) = parent
        {
            holder.clazz = parent.clazz.clone();
        } else {
            holder.clazz = basic_parameter.clazz.clone();
        }

        holder.global_configuration = GlobalConfiguration::new();
        holder.global_configuration.auto_trim = basic_parameter
            .auto_trim
            .or_else(|| parent.map(|parent| parent.global_configuration.auto_trim))
            .unwrap_or(true);
        holder.global_configuration.use1904windowing = basic_parameter
            .use1904windowing
            .or_else(|| parent.map(|parent| parent.global_configuration.use1904windowing))
            .unwrap_or(false);
        holder.global_configuration.locale = basic_parameter
            .locale
            .clone()
            .or_else(|| parent.map(|parent| parent.global_configuration.locale.clone()))
            .unwrap_or_else(|| "default".to_owned());
        holder.global_configuration.use_scientific_format = basic_parameter
            .use_scientific_format
            .or_else(|| parent.map(|parent| parent.global_configuration.use_scientific_format))
            .unwrap_or(false);
        holder.global_configuration.filed_cache_location = basic_parameter
            .filed_cache_location
            .or_else(|| parent.map(|parent| parent.global_configuration.filed_cache_location))
            .unwrap_or(CacheLocation::ThreadLocal);

        holder
    }

    /// Returns the dynamic header rows. (Java `getHead()`)
    #[must_use]
    pub fn head(&self) -> Option<&[Vec<String>]> {
        self.head.as_deref()
    }

    /// Returns the model type name. (Java `getClazz()`)
    #[must_use]
    pub fn clazz(&self) -> Option<&str> {
        self.clazz.as_deref()
    }
}

impl super::configuration_holder::MetadataHolder for AbstractHolder {
    fn holder_type(&self) -> HolderEnum {
        self.holder_type
    }
}

impl ConfigurationHolder for AbstractHolder {
    fn is_new(&self) -> bool {
        self.new_initialization
    }

    fn global_configuration(&self) -> &GlobalConfiguration {
        &self.global_configuration
    }

    fn converter_map(&self) -> &ConverterRegistry {
        &self.converter_map
    }
}
