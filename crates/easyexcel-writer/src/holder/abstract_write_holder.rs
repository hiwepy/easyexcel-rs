//! Mirrors Java `com.alibaba.excel.write.metadata.holder.AbstractWriteHolder`.

use std::collections::HashSet;

use easyexcel_core::ConverterRegistry;
use easyexcel_core::ExcelCellStyle;
use easyexcel_core::ExcelFontStyle;
use easyexcel_core::ExcelWriteMetadata;
use easyexcel_core::converter::default_converter_loader::load_default_write_converter;

use crate::metadata::WriteBasicParameter;
use crate::{ExcelWriteHeadProperty, WriteHolder};

/// Mirrors Java `AbstractWriteHolder extends AbstractHolder implements WriteHolder`.
///
/// The Java side carries resolved nullable parameters inherited from the
/// parent holder. Rust keeps the same resolved state here; builders use
/// [`crate::WriteOptions`] for the live backend while handler-facing
/// compatibility APIs use this holder.
#[derive(Debug, Clone)]
pub struct AbstractWriteHolder {
    /// Mirrors `AbstractWriteHolder.needHead`.
    pub need_head: bool,
    /// Mirrors `AbstractWriteHolder.relativeHeadRowIndex`.
    pub relative_head_row_index: i32,
    /// Mirrors `AbstractWriteHolder.useDefaultStyle`.
    pub use_default_style: bool,
    /// Mirrors `AbstractWriteHolder.automaticMergeHead`.
    pub automatic_merge_head: bool,
    /// Mirrors `AbstractWriteHolder.excelWriteHeadProperty`.
    pub excel_write_head_property: ExcelWriteHeadProperty,
    /// Mirrors `AbstractWriteHolder.headStyle`.
    pub head_style: Option<ExcelCellStyle>,
    /// Mirrors `AbstractWriteHolder.contentStyle`.
    pub content_style: Option<ExcelCellStyle>,
    /// Mirrors `AbstractWriteHolder.headFontStyle`.
    pub head_font_style: Option<ExcelFontStyle>,
    /// Mirrors `AbstractWriteHolder.contentFontStyle`.
    pub content_font_style: Option<ExcelFontStyle>,
    /// Mirrors `AbstractWriteHolder.excludeColumnIndexes`.
    pub exclude_column_indexes: Option<HashSet<usize>>,
    /// Mirrors `AbstractWriteHolder.excludeColumnFieldNames`.
    pub exclude_column_field_names: Option<HashSet<String>>,
    /// Mirrors `AbstractWriteHolder.includeColumnIndexes`.
    pub include_column_indexes: Option<HashSet<usize>>,
    /// Mirrors `AbstractWriteHolder.includeColumnFieldNames`.
    pub include_column_field_names: Option<HashSet<String>>,
    /// Mirrors `AbstractWriteHolder.orderByIncludeColumn`.
    pub order_by_include_column: bool,
    /// Mirrors `AbstractHolder.converterMap`.
    pub converter_map: ConverterRegistry,
}

impl Default for AbstractWriteHolder {
    fn default() -> Self {
        Self {
            need_head: true,
            relative_head_row_index: 0,
            use_default_style: true,
            automatic_merge_head: true,
            excel_write_head_property: ExcelWriteHeadProperty::new(),
            head_style: None,
            content_style: None,
            head_font_style: None,
            content_font_style: None,
            exclude_column_indexes: None,
            exclude_column_field_names: None,
            include_column_indexes: None,
            include_column_field_names: None,
            order_by_include_column: false,
            converter_map: load_default_write_converter(),
        }
    }
}

impl AbstractWriteHolder {
    /// Resolves Java nullable write parameters against an optional parent.
    ///
    /// A missing collection inherits the parent collection, while an explicit
    /// empty collection clears it. This distinction is required by Java
    /// `AbstractWriteHolder(WriteBasicParameter, parent)`.
    #[must_use]
    pub fn from_parameter(
        parameter: &WriteBasicParameter,
        parent: Option<&AbstractWriteHolder>,
    ) -> Self {
        let defaults = Self::default();
        Self {
            need_head: parameter
                .need_head
                .or_else(|| parent.map(|holder| holder.need_head))
                .unwrap_or(defaults.need_head),
            relative_head_row_index: parameter
                .relative_head_row_index
                .or_else(|| parent.map(|holder| holder.relative_head_row_index))
                .unwrap_or(defaults.relative_head_row_index),
            use_default_style: parameter
                .use_default_style
                .or_else(|| parent.map(|holder| holder.use_default_style))
                .unwrap_or(defaults.use_default_style),
            automatic_merge_head: parameter
                .automatic_merge_head
                .or_else(|| parent.map(|holder| holder.automatic_merge_head))
                .unwrap_or(defaults.automatic_merge_head),
            exclude_column_indexes: resolve_set(
                parameter.exclude_column_indexes.as_ref(),
                parent.and_then(|holder| holder.exclude_column_indexes.as_ref()),
            ),
            exclude_column_field_names: resolve_set(
                parameter.exclude_column_field_names.as_ref(),
                parent.and_then(|holder| holder.exclude_column_field_names.as_ref()),
            ),
            include_column_indexes: resolve_set(
                parameter.include_column_indexes.as_ref(),
                parent.and_then(|holder| holder.include_column_indexes.as_ref()),
            ),
            include_column_field_names: resolve_set(
                parameter.include_column_field_names.as_ref(),
                parent.and_then(|holder| holder.include_column_field_names.as_ref()),
            ),
            order_by_include_column: parameter
                .order_by_include_column
                .or_else(|| parent.map(|holder| holder.order_by_include_column))
                .unwrap_or(defaults.order_by_include_column),
            excel_write_head_property: ExcelWriteHeadProperty::new(),
            head_style: parent.and_then(|holder| holder.head_style),
            content_style: parent.and_then(|holder| holder.content_style),
            head_font_style: parent.and_then(|holder| holder.head_font_style),
            content_font_style: parent.and_then(|holder| holder.content_font_style),
            converter_map: parent
                .map_or_else(load_default_write_converter, |holder| {
                    holder.converter_map.clone()
                })
                .merged_with(&parameter.converters),
        }
    }

    /// Returns the effective converter map inherited by this holder.
    /// (Java `ConfigurationHolder.converterMap()`)
    #[must_use]
    pub const fn converter_map(&self) -> &ConverterRegistry {
        &self.converter_map
    }

    /// Replaces the resolved head property carried by this holder.
    ///
    /// Java creates this property during every holder constructor. Rust
    /// builders can resolve schema and dynamic-head information later, so the
    /// assignment is explicit rather than hidden behind a metadata placeholder.
    pub fn set_excel_write_head_property(&mut self, property: ExcelWriteHeadProperty) {
        self.excel_write_head_property = property;
    }

    /// Resolves a raw dynamic/class head into this holder. (Java constructor)
    pub fn resolve_head(
        &mut self,
        head_clazz: Option<String>,
        head: Option<Vec<Vec<String>>>,
        metadata: ExcelWriteMetadata,
    ) {
        self.excel_write_head_property =
            ExcelWriteHeadProperty::from_head(None, head_clazz, head, metadata);
    }
}

impl WriteHolder for AbstractWriteHolder {
    fn excel_write_head_property(&self) -> &ExcelWriteHeadProperty {
        &self.excel_write_head_property
    }

    fn ignore(&self, field_name: Option<&str>, column_index: Option<usize>) -> bool {
        if let Some(field_name) = field_name {
            if self
                .include_column_field_names
                .as_ref()
                .is_some_and(|names| !names.contains(field_name))
            {
                return true;
            }
            if self
                .exclude_column_field_names
                .as_ref()
                .is_some_and(|names| names.contains(field_name))
            {
                return true;
            }
        }
        if let Some(column_index) = column_index {
            if self
                .include_column_indexes
                .as_ref()
                .is_some_and(|indexes| !indexes.contains(&column_index))
            {
                return true;
            }
            if self
                .exclude_column_indexes
                .as_ref()
                .is_some_and(|indexes| indexes.contains(&column_index))
            {
                return true;
            }
        }
        false
    }

    fn need_head(&self) -> bool {
        self.need_head
    }

    fn relative_head_row_index(&self) -> i32 {
        self.relative_head_row_index
    }

    fn automatic_merge_head(&self) -> bool {
        self.automatic_merge_head
    }

    fn order_by_include_column(&self) -> bool {
        self.order_by_include_column
    }

    fn include_column_indexes(&self) -> Option<&HashSet<usize>> {
        self.include_column_indexes.as_ref()
    }

    fn include_column_field_names(&self) -> Option<&HashSet<String>> {
        self.include_column_field_names.as_ref()
    }

    fn exclude_column_indexes(&self) -> Option<&HashSet<usize>> {
        self.exclude_column_indexes.as_ref()
    }

    fn exclude_column_field_names(&self) -> Option<&HashSet<String>> {
        self.exclude_column_field_names.as_ref()
    }
}

fn resolve_set<T>(own: Option<&Vec<T>>, parent: Option<&HashSet<T>>) -> Option<HashSet<T>>
where
    T: Clone + Eq + std::hash::Hash,
{
    own.map(|values| values.iter().cloned().collect())
        .or_else(|| parent.cloned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[derive(Clone, Copy)]
    struct PrefixConverter(&'static str);

    impl easyexcel_core::Converter<String> for PrefixConverter {
        fn convert_to_excel_data(
            &self,
            context: &easyexcel_core::WriteConverterContext<'_, String>,
        ) -> easyexcel_core::Result<easyexcel_core::WriteCellData> {
            Ok(easyexcel_core::WriteCellData::from_string(format!(
                "{}:{}",
                self.0,
                context.value()
            )))
        }
    }

    fn convert_string(holder: &AbstractWriteHolder, value: &str) -> String {
        holder
            .converter_map()
            .convert_to_excel_data(
                &value.to_owned(),
                &easyexcel_core::ExcelColumn::new("value", "Value", Some(0), 0, None),
                &easyexcel_core::ConvertContext {
                    sheet_name: "Data".to_owned(),
                    row_index: 1,
                    column_index: Some(0),
                    field: "value",
                    format: None,
                    use_1904_windowing: false,
                },
            )
            .expect("converter succeeds")
            .expect("converter registered")
            .value()
            .as_text()
    }

    #[test]
    fn java_root_defaults_and_parent_inheritance_are_resolved() {
        let root = AbstractWriteHolder::from_parameter(&WriteBasicParameter::default(), None);
        assert!(root.need_head);
        assert!(root.use_default_style);
        assert!(root.automatic_merge_head);
        assert_eq!(root.relative_head_row_index, 0);
        assert!(!root.order_by_include_column);

        let parent = AbstractWriteHolder::from_parameter(
            &WriteBasicParameter {
                need_head: Some(false),
                include_column_indexes: Some(vec![1, 3]),
                exclude_column_field_names: Some(vec!["secret".to_owned()]),
                order_by_include_column: Some(true),
                ..WriteBasicParameter::default()
            },
            None,
        );
        let child =
            AbstractWriteHolder::from_parameter(&WriteBasicParameter::default(), Some(&parent));
        assert!(!child.need_head);
        assert_eq!(child.include_column_indexes, parent.include_column_indexes);
        assert_eq!(
            child.exclude_column_field_names,
            parent.exclude_column_field_names
        );
        assert!(child.order_by_include_column);
    }

    #[test]
    fn explicit_empty_collection_clears_parent_and_ignore_matches_java() {
        let parent = AbstractWriteHolder::from_parameter(
            &WriteBasicParameter {
                include_column_indexes: Some(vec![1, 3]),
                include_column_field_names: Some(vec!["name".to_owned(), "age".to_owned()]),
                exclude_column_field_names: Some(vec!["age".to_owned()]),
                ..WriteBasicParameter::default()
            },
            None,
        );
        assert!(!parent.ignore(Some("name"), Some(1)));
        assert!(parent.ignore(Some("other"), Some(1)));
        assert!(parent.ignore(Some("age"), Some(1)));
        assert!(parent.ignore(Some("name"), Some(2)));

        let child = AbstractWriteHolder::from_parameter(
            &WriteBasicParameter {
                include_column_indexes: Some(Vec::new()),
                include_column_field_names: Some(Vec::new()),
                exclude_column_field_names: Some(Vec::new()),
                ..WriteBasicParameter::default()
            },
            Some(&parent),
        );
        assert!(child.ignore(Some("name"), Some(1)));
        assert_eq!(child.include_column_indexes, Some(HashSet::new()));
        assert_eq!(child.exclude_column_field_names, Some(HashSet::new()));
    }

    #[test]
    fn converter_map_clones_parent_and_applies_child_override() {
        let mut parent_parameter = WriteBasicParameter::default();
        parent_parameter
            .converters
            .register::<String, _>(PrefixConverter("parent"));
        let parent = AbstractWriteHolder::from_parameter(&parent_parameter, None);

        let inherited =
            AbstractWriteHolder::from_parameter(&WriteBasicParameter::default(), Some(&parent));
        assert_eq!(convert_string(&inherited, "value"), "parent:value");

        let mut child_parameter = WriteBasicParameter::default();
        child_parameter
            .converters
            .register::<String, _>(PrefixConverter("child"));
        let child = AbstractWriteHolder::from_parameter(&child_parameter, Some(&parent));
        assert_eq!(convert_string(&child, "value"), "child:value");
        assert_eq!(convert_string(&parent, "value"), "parent:value");
    }

    #[test]
    fn holder_exposes_real_head_property_and_complete_selection_surface() {
        let mut holder = AbstractWriteHolder::from_parameter(
            &WriteBasicParameter {
                include_column_indexes: Some(vec![2, 4]),
                include_column_field_names: Some(vec!["name".to_owned()]),
                exclude_column_indexes: Some(vec![7]),
                exclude_column_field_names: Some(vec!["secret".to_owned()]),
                order_by_include_column: Some(true),
                ..WriteBasicParameter::default()
            },
            None,
        );
        holder.resolve_head(
            Some("DemoData".to_owned()),
            Some(vec![vec!["用户".to_owned(), "姓名".to_owned()]]),
            ExcelWriteMetadata::new().head_row_height(26),
        );

        let contract: &dyn WriteHolder = &holder;
        assert_eq!(
            contract.excel_write_head_property().head_clazz(),
            Some("DemoData")
        );
        assert_eq!(contract.excel_write_head_property().head_row_number(), 2);
        assert_eq!(
            contract
                .excel_write_head_property()
                .head_row_height_property()
                .map(easyexcel_core::metadata::RowHeightProperty::height),
            Some(26)
        );
        assert!(contract.order_by_include_column());
        assert_eq!(
            contract.include_column_indexes(),
            Some(&HashSet::from([2, 4]))
        );
        assert_eq!(
            contract.include_column_field_names(),
            Some(&HashSet::from(["name".to_owned()]))
        );
        assert_eq!(contract.exclude_column_indexes(), Some(&HashSet::from([7])));
        assert_eq!(
            contract.exclude_column_field_names(),
            Some(&HashSet::from(["secret".to_owned()]))
        );
    }
}
