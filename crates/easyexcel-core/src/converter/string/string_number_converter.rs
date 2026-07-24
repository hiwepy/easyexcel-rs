//! Mirrors Java `com.alibaba.excel.converters.string.StringNumberConverter`.

/// Mirrors Java `StringNumberConverter`.
#[derive(Debug, Clone, Copy, Default)]
pub struct StringNumberConverter;

impl crate::Converter<String> for StringNumberConverter {
    fn support_excel_type(&self) -> crate::CellDataType {
        crate::CellDataType::Number
    }

    fn convert_to_rust_data(
        &self,
        context: &crate::ReadConverterContext<'_>,
    ) -> Result<String, crate::ExcelError> {
        use bigdecimal::BigDecimal;

        let cell = context.cell().unwrap_or(&crate::CellValue::Empty);
        let decimal = if let Some(value) = context.decimal_value() {
            value.clone()
        } else {
            match cell {
                crate::CellValue::Decimal(value) => value.clone(),
                crate::CellValue::Int(value) => BigDecimal::from(*value),
                crate::CellValue::Float(value) if value.is_finite() => value
                    .to_string()
                    .parse()
                    .map_err(|_| context.convert_context().invalid(cell, "String"))?,
                _ => return Err(context.convert_context().invalid(cell, "String")),
            }
        };

        if let Some(pattern) = context
            .column()
            .format
            .or(context.convert_context().format)
            .filter(|pattern| !pattern.is_empty())
        {
            if crate::util::date_utils::is_internal_date_format(pattern) {
                return crate::converter::date_support::format_number_as_datetime_string(
                    context, pattern,
                );
            }
            return crate::util::number_utils::format_decimal(
                &decimal,
                decimal < BigDecimal::from(0),
                Some(pattern),
                context.column().number_rounding_mode.unwrap_or_default(),
            );
        }

        if let Some(display) = context.display_value() {
            return Ok(display.to_owned());
        }
        Ok(decimal.to_plain_string())
    }

    fn convert_to_excel_data(
        &self,
        context: &crate::WriteConverterContext<'_, String>,
    ) -> Result<crate::WriteCellData, crate::ExcelError> {
        let value = context
            .value()
            .parse::<bigdecimal::BigDecimal>()
            .map_err(|_| {
                crate::ExcelError::Format(format!("invalid BigDecimal value {:?}", context.value()))
            })?;
        Ok(crate::WriteCellData::new(crate::CellValue::Decimal(value)))
    }
}

#[cfg(test)]
mod tests {
    use std::collections::{HashMap, HashSet};
    use std::sync::Arc;

    use bigdecimal::BigDecimal;

    use super::*;
    use crate::{
        CellValue, ConvertContext, Converter, ExcelColumn, ReadConverterContext, RowData,
        WriteConverterContext,
    };

    fn convert_context(format: Option<&'static str>) -> ConvertContext {
        ConvertContext {
            sheet_name: "Data".to_owned(),
            row_index: 1,
            column_index: Some(0),
            field: "value",
            format,
            use_1904_windowing: false,
        }
    }

    #[test]
    fn reads_explicit_formats_and_source_display_without_losing_decimal_scale() {
        let converter = StringNumberConverter;
        let number = CellValue::Float(0.125);
        let number_column = ExcelColumn::new("value", "Value", Some(0), 0, Some("0.0%"));
        let context = convert_context(Some("0.0%"));
        assert_eq!(
            converter
                .convert_to_rust_data(&ReadConverterContext::new(
                    Some(&number),
                    &number_column,
                    &context,
                ))
                .expect("number format"),
            "12.5%"
        );

        let headers = Arc::new(HashMap::from([("Value".to_owned(), 0)]));
        let row = RowData::new("Data", 1, vec![number], headers)
            .with_display_values(HashMap::from([(0, "12.50%".to_owned())]))
            .with_decimal_values(HashMap::from([(
                0,
                "0.1250".parse::<BigDecimal>().expect("decimal"),
            )]))
            .with_present_columns(HashSet::from([0]));
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        let context = row.convert_context(&column);
        assert_eq!(
            converter
                .convert_to_rust_data(&ReadConverterContext::with_cell_metadata(
                    row.cell(&column),
                    row.formula(&column),
                    row.display_value(&column),
                    row.decimal_value(&column),
                    &column,
                    &context,
                ))
                .expect("source display"),
            "12.50%"
        );
    }

    #[test]
    fn writes_strict_big_decimal_number_cell() {
        let converter = StringNumberConverter;
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        let context = convert_context(None);
        let value = "12.30".to_owned();
        assert_eq!(
            converter
                .convert_to_excel_data(&WriteConverterContext::new(&value, &column, &context,))
                .expect("decimal")
                .value(),
            &CellValue::Decimal("12.30".parse().expect("decimal"))
        );
        let invalid = " 12.30 ".to_owned();
        assert!(
            converter
                .convert_to_excel_data(&WriteConverterContext::new(&invalid, &column, &context,))
                .is_err()
        );
    }
}
