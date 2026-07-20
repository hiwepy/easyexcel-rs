//! Mirrors Java `com.alibaba.excel.metadata.ConfigurationHolder` and `Holder`.

use crate::Holder as HolderEnum;
use crate::ConverterRegistry;

use super::global_configuration::GlobalConfiguration;

/// Java `Holder` interface contract.
///
/// The core crate already exports [`HolderEnum`] for Java `HolderEnum`. This
/// trait mirrors Java `Holder.holderType()` without colliding with that enum
/// name.
pub trait MetadataHolder {
    /// Returns the holder scope. (Java `holderType()`)
    fn holder_type(&self) -> HolderEnum;
}

/// Read/write holder configuration contract.
///
/// Rust port of Java `ConfigurationHolder extends Holder`.
pub trait ConfigurationHolder: MetadataHolder {
    /// Returns whether the holder was freshly initialized. (Java `isNew()`)
    fn is_new(&self) -> bool;

    /// Returns the global configuration. (Java `globalConfiguration()`)
    fn global_configuration(&self) -> &GlobalConfiguration;

    /// Returns the active converter registry. (Java `converterMap()`)
    fn converter_map(&self) -> &ConverterRegistry;
}
