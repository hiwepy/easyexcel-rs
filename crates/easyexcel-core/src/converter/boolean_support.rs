use bigdecimal::BigDecimal;
use num_bigint::BigInt;

use crate::{CellValue, ExcelError, ReadConverterContext, WriteCellData, WriteConverterContext};

pub(crate) trait BooleanScalar: Sized {
    fn from_boolean(value: bool) -> Self;
    fn is_one(&self) -> bool;
}

macro_rules! impl_boolean_scalar {
    ($($target:ty),+ $(,)?) => {
        $(
            impl BooleanScalar for $target {
                fn from_boolean(value: bool) -> Self {
                    if value { 1 as $target } else { 0 as $target }
                }

                fn is_one(&self) -> bool {
                    *self == 1 as $target
                }
            }
        )+
    };
}

impl_boolean_scalar!(i8, i16, i32, i64, f32, f64);

impl BooleanScalar for BigDecimal {
    fn from_boolean(value: bool) -> Self {
        Self::from(i32::from(value))
    }

    fn is_one(&self) -> bool {
        self == &Self::from(1)
    }
}

impl BooleanScalar for BigInt {
    fn from_boolean(value: bool) -> Self {
        Self::from(i32::from(value))
    }

    fn is_one(&self) -> bool {
        self == &Self::from(1)
    }
}

pub(crate) fn read_boolean_scalar<T>(context: &ReadConverterContext<'_>) -> Result<T, ExcelError>
where
    T: BooleanScalar,
{
    match context.cell() {
        Some(CellValue::Bool(value)) => Ok(T::from_boolean(*value)),
        Some(value) => Err(context.convert_context().invalid(value, "boolean scalar")),
        None => Err(context
            .convert_context()
            .invalid(&CellValue::Empty, "boolean scalar")),
    }
}

pub(crate) fn write_scalar_boolean<T>(context: &WriteConverterContext<'_, T>) -> WriteCellData
where
    T: BooleanScalar,
{
    WriteCellData::new(CellValue::Bool(context.value().is_one()))
}

pub(crate) fn read_boolean(context: &ReadConverterContext<'_>) -> Result<bool, ExcelError> {
    match context.cell() {
        Some(CellValue::Bool(value)) => Ok(*value),
        Some(value) => Err(context.convert_context().invalid(value, "bool")),
        None => Err(context.convert_context().invalid(&CellValue::Empty, "bool")),
    }
}

pub(crate) fn read_number_boolean(context: &ReadConverterContext<'_>) -> Result<bool, ExcelError> {
    match context.cell() {
        Some(CellValue::Int(value)) => Ok(*value == 1),
        Some(CellValue::Float(value)) => Ok(*value == 1.0),
        Some(CellValue::Decimal(value)) => Ok(value == &BigDecimal::from(1)),
        Some(value) => Err(context.convert_context().invalid(value, "bool")),
        None => Err(context.convert_context().invalid(&CellValue::Empty, "bool")),
    }
}

pub(crate) fn write_boolean_number(context: &WriteConverterContext<'_, bool>) -> WriteCellData {
    WriteCellData::new(CellValue::Decimal(BigDecimal::from(i32::from(
        *context.value(),
    ))))
}

pub(crate) fn read_string_boolean(context: &ReadConverterContext<'_>) -> Result<bool, ExcelError> {
    match context.cell() {
        Some(CellValue::String(value)) => Ok(value.eq_ignore_ascii_case("true")),
        Some(value) => Err(context.convert_context().invalid(value, "bool")),
        None => Err(context.convert_context().invalid(&CellValue::Empty, "bool")),
    }
}

pub(crate) fn write_boolean_string(context: &WriteConverterContext<'_, bool>) -> WriteCellData {
    WriteCellData::from_string(context.value().to_string())
}

pub(crate) fn read_boolean_string_value(
    context: &ReadConverterContext<'_>,
) -> Result<String, ExcelError> {
    read_boolean(context).map(|value| value.to_string())
}

pub(crate) fn write_string_boolean(context: &WriteConverterContext<'_, String>) -> WriteCellData {
    WriteCellData::new(CellValue::Bool(
        context.value().eq_ignore_ascii_case("true"),
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::converter::bigdecimal::big_decimal_boolean_converter::BigDecimalBooleanConverter;
    use crate::converter::biginteger::big_integer_boolean_converter::BigIntegerBooleanConverter;
    use crate::converter::booleanconverter::boolean_boolean_converter::BooleanBooleanConverter;
    use crate::converter::booleanconverter::boolean_number_converter::BooleanNumberConverter;
    use crate::converter::booleanconverter::boolean_string_converter::BooleanStringConverter;
    use crate::converter::byteconverter::byte_boolean_converter::ByteBooleanConverter;
    use crate::converter::doubleconverter::double_boolean_converter::DoubleBooleanConverter;
    use crate::converter::floatconverter::float_boolean_converter::FloatBooleanConverter;
    use crate::converter::integer::integer_boolean_converter::IntegerBooleanConverter;
    use crate::converter::longconverter::long_boolean_converter::LongBooleanConverter;
    use crate::converter::shortconverter::short_boolean_converter::ShortBooleanConverter;
    use crate::converter::string::string_boolean_converter::StringBooleanConverter;
    use crate::{
        CellDataType, ConvertContext, Converter, ExcelColumn, ReadConverterContext,
        WriteConverterContext,
    };

    const COLUMN: ExcelColumn = ExcelColumn::new("value", "Value", Some(0), 0, None);

    fn context() -> ConvertContext {
        ConvertContext {
            sheet_name: "Sheet1".to_owned(),
            row_index: 1,
            column_index: Some(0),
            field: "value",
            format: None,
            use_1904_windowing: false,
        }
    }

    #[test]
    fn scalar_boolean_converters_use_exact_java_zero_one_rules() {
        let context = context();
        let true_cell = CellValue::Bool(true);
        let read = ReadConverterContext::new(Some(&true_cell), &COLUMN, &context);
        assert_eq!(
            BigDecimalBooleanConverter.convert_to_rust_data(&read),
            Ok(BigDecimal::from(1))
        );
        assert_eq!(
            BigIntegerBooleanConverter.convert_to_rust_data(&read),
            Ok(BigInt::from(1))
        );
        assert_eq!(ByteBooleanConverter.convert_to_rust_data(&read), Ok(1));
        assert_eq!(ShortBooleanConverter.convert_to_rust_data(&read), Ok(1));
        assert_eq!(IntegerBooleanConverter.convert_to_rust_data(&read), Ok(1));
        assert_eq!(LongBooleanConverter.convert_to_rust_data(&read), Ok(1));
        assert_eq!(FloatBooleanConverter.convert_to_rust_data(&read), Ok(1.0));
        assert_eq!(DoubleBooleanConverter.convert_to_rust_data(&read), Ok(1.0));

        let one = 1_i32;
        let two = 2_i32;
        assert_eq!(
            IntegerBooleanConverter
                .convert_to_excel_data(&WriteConverterContext::new(&one, &COLUMN, &context))
                .unwrap()
                .value(),
            &CellValue::Bool(true)
        );
        assert_eq!(
            IntegerBooleanConverter
                .convert_to_excel_data(&WriteConverterContext::new(&two, &COLUMN, &context))
                .unwrap()
                .value(),
            &CellValue::Bool(false)
        );
    }

    #[test]
    fn boolean_number_and_string_converters_match_java_value_of() {
        let context = context();
        for (cell, expected) in [
            (CellValue::Int(1), true),
            (CellValue::Float(1.0), true),
            (CellValue::Decimal(BigDecimal::from(1)), true),
            (CellValue::Int(0), false),
            (CellValue::Int(2), false),
            (CellValue::Float(-1.0), false),
        ] {
            let read = ReadConverterContext::new(Some(&cell), &COLUMN, &context);
            assert_eq!(
                BooleanNumberConverter.convert_to_rust_data(&read),
                Ok(expected)
            );
        }

        for (text, expected) in [
            ("true", true),
            ("TrUe", true),
            ("false", false),
            ("1", false),
            (" true ", false),
            ("anything", false),
        ] {
            let cell = CellValue::String(text.to_owned());
            let read = ReadConverterContext::new(Some(&cell), &COLUMN, &context);
            assert_eq!(
                BooleanStringConverter.convert_to_rust_data(&read),
                Ok(expected)
            );
        }

        for value in [true, false] {
            let write = WriteConverterContext::new(&value, &COLUMN, &context);
            assert_eq!(
                BooleanBooleanConverter
                    .convert_to_excel_data(&write)
                    .unwrap()
                    .value(),
                &CellValue::Bool(value)
            );
            assert_eq!(
                BooleanStringConverter
                    .convert_to_excel_data(&write)
                    .unwrap()
                    .value(),
                &CellValue::String(value.to_string())
            );
            assert_eq!(
                BooleanNumberConverter
                    .convert_to_excel_data(&write)
                    .unwrap()
                    .value(),
                &CellValue::Decimal(BigDecimal::from(i32::from(value)))
            );
        }
    }

    #[test]
    fn string_boolean_converter_is_bidirectional_and_registered_by_boolean_source() {
        let context = context();
        let true_cell = CellValue::Bool(true);
        let read = ReadConverterContext::new(Some(&true_cell), &COLUMN, &context);
        assert_eq!(
            StringBooleanConverter.convert_to_rust_data(&read),
            Ok("true".to_owned())
        );
        for (text, expected) in [("TRUE", true), ("1", false), ("yes", false)] {
            let text = text.to_owned();
            let write = WriteConverterContext::new(&text, &COLUMN, &context);
            assert_eq!(
                StringBooleanConverter
                    .convert_to_excel_data(&write)
                    .unwrap()
                    .value(),
                &CellValue::Bool(expected)
            );
        }
        assert_eq!(
            <StringBooleanConverter as Converter<String>>::support_excel_type(
                &StringBooleanConverter
            ),
            CellDataType::Boolean
        );
    }
}
