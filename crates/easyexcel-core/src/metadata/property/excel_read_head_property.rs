//! Mirrors Java `com.alibaba.excel.read.metadata.property.ExcelReadHeadProperty`.

use super::super::configuration_holder::ConfigurationHolder;
use super::excel_head_property::ExcelHeadProperty;

/// Read-side header metadata.
///
/// Rust port of Java `ExcelReadHeadProperty extends ExcelHeadProperty`.
#[derive(Debug, Clone, PartialEq)]
pub struct ExcelReadHeadProperty(ExcelHeadProperty);

impl ExcelReadHeadProperty {
    /// Creates read-side header metadata. (Java constructor)
    #[must_use]
    pub fn new(
        configuration_holder: Option<&dyn ConfigurationHolder>,
        head_clazz: Option<String>,
        head: Option<Vec<Vec<String>>>,
    ) -> Self {
        let property = if let Some(head_clazz) = head_clazz {
            ExcelHeadProperty::for_class(configuration_holder, head_clazz, head)
        } else {
            ExcelHeadProperty::new(configuration_holder, head)
        };
        Self(property)
    }

    /// Returns the underlying header property. (Java inherited getters)
    #[must_use]
    pub fn inner(&self) -> &ExcelHeadProperty {
        &self.0
    }

    /// Returns whether any header is configured. (Java `hasHead()`)
    #[must_use]
    pub fn has_head(&self) -> bool {
        self.0.has_head()
    }

    /// Returns the header map. (Java `getHeadMap()`)
    #[must_use]
    pub fn head_map(&self) -> &std::collections::BTreeMap<i32, super::super::head::Head> {
        self.0.head_map()
    }
}

impl std::ops::Deref for ExcelReadHeadProperty {
    type Target = ExcelHeadProperty;

    fn deref(&self) -> &Self::Target {
        &self.0
    }
}
