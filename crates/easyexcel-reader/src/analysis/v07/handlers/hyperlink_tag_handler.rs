//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.HyperlinkTagHandler`.

use std::collections::HashMap;

use easyexcel_core::constant::excel_xml_constants::{
    ATTRIBUTE_LOCATION, ATTRIBUTE_REF, ATTRIBUTE_RID,
};
use easyexcel_core::{CellExtra, CellExtraType, ExcelError, Result};

use super::merge_cell_tag_handler::cell_extra_from_ref;
use super::xlsx_tag_handler::XlsxTagHandler;

/// Mirrors Java `HyperlinkTagHandler`.
#[derive(Debug, Default)]
pub struct HyperlinkTagHandler {
    /// Whether hyperlink extras are enabled. (Java `support`)
    pub enabled: bool,
    /// Last parsed hyperlink extra.
    pub last_extra: Option<CellExtra>,
}

impl HyperlinkTagHandler {
    /// Creates a handler; `enabled` mirrors Java `support(XlsxReadContext)`.
    #[must_use]
    pub fn new(enabled: bool) -> Self {
        Self {
            enabled,
            last_extra: None,
        }
    }

    /// Java `HyperlinkTagHandler.startElement`.
    ///
    /// `resolve_r_id` maps `r:id` → target URI (Java `PackageRelationshipCollection`).
    pub fn start_hyperlink(
        &mut self,
        attrs: &HashMap<String, String>,
        resolve_r_id: &dyn Fn(&str) -> Option<String>,
    ) -> Result<()> {
        if !self.enabled {
            return Ok(());
        }
        let Some(reference) = attrs.get(ATTRIBUTE_REF) else {
            return Ok(());
        };
        if reference.is_empty() {
            return Ok(());
        }
        if let Some(location) = attrs.get(ATTRIBUTE_LOCATION) {
            self.last_extra = Some(cell_extra_from_ref(
                CellExtraType::Hyperlink,
                Some(location.clone()),
                reference,
            )?);
            return Ok(());
        }
        // Java `Attributes.get("r:id")`; `quick_xml` local-name strips the
        // prefix so worksheet SAX attrs arrive as plain `"id"`.
        let r_id = attrs
            .get(ATTRIBUTE_RID)
            .or_else(|| attrs.get("id"))
            .map(String::as_str);
        if let Some(r_id) = r_id
            && let Some(uri) = resolve_r_id(r_id)
        {
            self.last_extra = Some(cell_extra_from_ref(
                CellExtraType::Hyperlink,
                Some(uri),
                reference,
            )?);
        }
        Ok(())
    }

    /// Strict variant used by `xlsx_rows::parse_worksheet_extras`.
    ///
    /// Missing / empty `ref`, missing `id`/`location`, and unresolved
    /// relationships all return [`ExcelError`] (historical Rust reader behaviour).
    pub fn start_hyperlink_required(
        &mut self,
        attrs: &HashMap<String, String>,
        resolve_r_id: &dyn Fn(&str) -> Result<String>,
    ) -> Result<()> {
        if !self.enabled {
            return Ok(());
        }
        let reference = attrs
            .get(ATTRIBUTE_REF)
            .filter(|value| !value.is_empty())
            .ok_or_else(|| ExcelError::Format("hyperlink ref is missing".to_owned()))?;
        if let Some(location) = attrs.get(ATTRIBUTE_LOCATION) {
            self.last_extra = Some(cell_extra_from_ref(
                CellExtraType::Hyperlink,
                Some(location.clone()),
                reference,
            )?);
            return Ok(());
        }
        let r_id = attrs
            .get(ATTRIBUTE_RID)
            .or_else(|| attrs.get("id"))
            .ok_or_else(|| ExcelError::Format("hyperlink id is missing".to_owned()))?;
        let uri = resolve_r_id(r_id)?;
        self.last_extra = Some(cell_extra_from_ref(
            CellExtraType::Hyperlink,
            Some(uri),
            reference,
        )?);
        Ok(())
    }
}

impl XlsxTagHandler for HyperlinkTagHandler {
    fn support(&self) -> bool {
        self.enabled
    }

    /// Java `HyperlinkTagHandler.startElement` (location-only; `r:id` needs a resolver).
    fn start_element(&mut self, name: &str, attrs: &str) {
        let local = name.rsplit(':').next().unwrap_or(name);
        if local != "hyperlink" {
            return;
        }
        let mut map = HashMap::new();
        for token in attrs.split_whitespace() {
            if let Some((key, value)) = token.split_once('=') {
                map.insert(key.to_owned(), value.to_owned());
            }
        }
        let _ = self.start_hyperlink(&map, &|_| None);
    }
}
