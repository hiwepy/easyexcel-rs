//! Unit tests for metadata abstractions and property facades.

use std::collections::BTreeMap;

use bigdecimal::RoundingMode;

use crate::metadata::{
    AbstractCell, AbstractHolder, AbstractParameterBuilder, AnalysisCell, BasicParameter,
    BasicParameterBuilder, Cell, CellRange, ConfigurationHolder, DateTimeFormatProperty,
    ExcelHeadProperty, ExcelReadHeadProperty, FieldCache, FieldWrapper, GlobalConfiguration,
    Head, MetadataHolder, NullObject, NumberFormatProperty,
};
use crate::{CacheLocation, HeadKind, Holder, WriteLastRow, WriteLastRowTypeEnum, WriteTemplateAnalysisCellType};

#[test]
fn null_object_is_zero_sized_marker() {
    assert_eq!(NullObject::new(), NullObject::default());
}

#[test]
fn cell_range_exposes_java_indices() {
    let range = CellRange::new(1, 3, 0, 2);
    assert_eq!(range.first_row(), 1);
    assert_eq!(range.last_row(), 3);
    assert_eq!(range.first_col(), 0);
    assert_eq!(range.last_col(), 2);
}

#[test]
fn abstract_cell_implements_cell_trait() {
    let cell = AbstractCell::with_indices(4, 5);
    assert_eq!(cell.row_index(), Some(4));
    assert_eq!(cell.column_index(), Some(5));
}

#[test]
fn field_cache_stores_sorted_and_index_maps() {
    let mut sorted = BTreeMap::new();
    sorted.insert(0, FieldWrapper::new("name", vec!["姓名".to_owned()]));
    let cache = FieldCache::new(sorted.clone(), sorted);
    assert_eq!(cache.sorted_field_map().len(), 1);
    assert_eq!(cache.index_field_map().len(), 1);
}

#[test]
fn head_rejects_empty_header_name() {
    let result = Head::new(0, None, vec!["".to_owned()], false, true);
    assert!(result.is_err());
}

#[test]
fn abstract_holder_inherits_parent_configuration() {
    let mut parent = AbstractHolder::new(Holder::Workbook);
    parent.global_configuration.auto_trim = false;
    parent.global_configuration.locale = "zh-CN".to_owned();

    let mut parameter = BasicParameter::new();
    parameter.auto_trim = None;
    parameter.locale = None;

    let holder = AbstractHolder::from_parameter(&parameter, Some(&parent), Holder::Sheet);
    assert!(!holder.global_configuration().auto_trim());
    assert_eq!(holder.global_configuration().locale(), "zh-CN");
    assert!(holder.is_new());
}

#[test]
fn basic_parameter_builder_supports_java_setters() {
    let mut builder = BasicParameterBuilder::new();
    builder
        .head(vec![vec!["姓名".to_owned()]])
        .head_class("DemoData")
        .register_converter("DemoConverter")
        .use1904windowing(true)
        .locale("en-US")
        .filed_cache_location(CacheLocation::Memory)
        .auto_trim(false);
    let parameter = builder.build();

    assert_eq!(parameter.head(), Some([vec!["姓名".to_owned()]].as_slice()));
    assert_eq!(parameter.clazz(), Some("DemoData"));
    assert_eq!(parameter.custom_converter_list(), ["DemoConverter"]);
    assert_eq!(parameter.use1904windowing, Some(true));
    assert_eq!(parameter.locale.as_deref(), Some("en-US"));
    assert_eq!(parameter.filed_cache_location, Some(CacheLocation::Memory));
    assert_eq!(parameter.auto_trim, Some(false));
}

#[test]
fn date_time_and_number_format_properties_build_from_annotations() {
    let date = DateTimeFormatProperty::build(Some("yyyy-MM-dd"), Some(true))
        .expect("date format");
    assert_eq!(date.format(), "yyyy-MM-dd");
    assert!(date.use1904windowing());

    let number = NumberFormatProperty::build(Some("0.00"), Some(RoundingMode::HalfUp))
        .expect("number format");
    assert_eq!(number.format(), "0.00");
    assert_eq!(number.rounding_mode(), RoundingMode::HalfUp);
}

#[test]
fn excel_read_head_property_wraps_string_head() {
    let head = ExcelReadHeadProperty::new(
        None,
        None,
        Some(vec![vec!["姓名".to_owned()], vec!["年龄".to_owned()]]),
    );
    assert!(head.has_head());
    assert_eq!(head.head_kind(), HeadKind::String);
    assert_eq!(head.head_map().len(), 2);
}

#[test]
fn excel_head_property_for_class_marks_class_kind() {
    let head = ExcelHeadProperty::for_class(None, "DemoData", None);
    assert_eq!(head.head_kind(), HeadKind::Class);
    assert_eq!(head.head_clazz(), Some("DemoData"));
}

#[test]
fn analysis_cell_tracks_template_placeholder() {
    let mut cell = AnalysisCell::new(2, 3);
    cell.variable_list.push("name".to_owned());
    cell.cell_type = WriteTemplateAnalysisCellType::Collection;
    assert_eq!(cell.row_index(), 2);
    assert_eq!(cell.column_index(), 3);
    assert_eq!(cell.cell_type(), WriteTemplateAnalysisCellType::Collection);
}

#[test]
fn write_last_row_type_enum_alias_matches_write_last_row() {
    let state: WriteLastRowTypeEnum = WriteLastRow::TemplateEmpty;
    assert_eq!(state, WriteLastRow::TemplateEmpty);
}

#[test]
fn global_configuration_defaults_match_java() {
    let config = GlobalConfiguration::new();
    assert!(config.auto_trim());
    assert!(!config.use1904windowing());
    assert_eq!(config.filed_cache_location(), CacheLocation::ThreadLocal);
}

#[test]
fn configuration_holder_is_implemented_by_abstract_holder() {
    let holder = AbstractHolder::new(Holder::Table);
    assert_eq!(holder.holder_type(), Holder::Table);
    assert!(holder.converter_map().is_empty());
}
