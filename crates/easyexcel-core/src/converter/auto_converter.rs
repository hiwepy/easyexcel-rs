//! Mirrors Java `com.alibaba.excel.converters.AutoConverter`.

use super::converter_trait::Converter;

/// Mirrors Java `AutoConverter implements Converter<Object>`.
///
/// This is deliberately a converter with only the default unsupported
/// operations. Java uses it as the `@ExcelProperty.converter` sentinel: when
/// it is present the runtime performs type-based converter lookup instead of
/// invoking this value.
#[derive(Debug, Clone, Copy, Default)]
pub struct AutoConverter;

impl<T> Converter<T> for AutoConverter {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{ConvertContext, ExcelColumn, ReadConverterContext};

    #[test]
    fn auto_converter_is_a_non_invokable_type_lookup_sentinel() {
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        let context = ConvertContext {
            sheet_name: "Data".to_owned(),
            row_index: 1,
            column_index: Some(0),
            field: "value",
            format: None,
            use_1904_windowing: false,
        };
        let error = <AutoConverter as Converter<String>>::convert_to_rust_data(
            &AutoConverter,
            &ReadConverterContext::new(None, &column, &context),
        )
        .expect_err("Java AutoConverter keeps Converter's unsupported defaults");
        assert!(matches!(error, crate::ExcelError::Unsupported(_)));
    }
}
