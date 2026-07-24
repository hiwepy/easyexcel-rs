//! Mirrors Java `com.alibaba.excel.converters.DefaultConverterLoader`.

use std::marker::PhantomData;
use std::path::PathBuf;

use bigdecimal::BigDecimal;
use chrono::{NaiveDate, NaiveDateTime};
use num_bigint::BigInt;
use url::Url;

use crate::InputStreamImageConverter;
use crate::converter::Converter;
use crate::converter::bigdecimal::big_decimal_boolean_converter::BigDecimalBooleanConverter;
use crate::converter::bigdecimal::big_decimal_number_converter::BigDecimalNumberConverter;
use crate::converter::bigdecimal::big_decimal_string_converter::BigDecimalStringConverter;
use crate::converter::biginteger::big_integer_boolean_converter::BigIntegerBooleanConverter;
use crate::converter::biginteger::big_integer_number_converter::BigIntegerNumberConverter;
use crate::converter::biginteger::big_integer_string_converter::BigIntegerStringConverter;
use crate::converter::booleanconverter::boolean_boolean_converter::BooleanBooleanConverter;
use crate::converter::booleanconverter::boolean_number_converter::BooleanNumberConverter;
use crate::converter::booleanconverter::boolean_string_converter::BooleanStringConverter;
use crate::converter::bytearray::boxing_byte_array_image_converter::BoxingByteArrayImageConverter;
use crate::converter::bytearray::byte_array_image_converter::ByteArrayImageConverter;
use crate::converter::byteconverter::byte_boolean_converter::ByteBooleanConverter;
use crate::converter::byteconverter::byte_number_converter::ByteNumberConverter;
use crate::converter::byteconverter::byte_string_converter::ByteStringConverter;
use crate::converter::date::date_date_converter::DateDateConverter;
use crate::converter::date::date_number_converter::DateNumberConverter;
use crate::converter::date::date_string_converter::DateStringConverter;
use crate::converter::doubleconverter::double_boolean_converter::DoubleBooleanConverter;
use crate::converter::doubleconverter::double_number_converter::DoubleNumberConverter;
use crate::converter::doubleconverter::double_string_converter::DoubleStringConverter;
use crate::converter::file::file_image_converter::FileImageConverter;
use crate::converter::floatconverter::float_boolean_converter::FloatBooleanConverter;
use crate::converter::floatconverter::float_number_converter::FloatNumberConverter;
use crate::converter::floatconverter::float_string_converter::FloatStringConverter;
use crate::converter::integer::integer_boolean_converter::IntegerBooleanConverter;
use crate::converter::integer::integer_number_converter::IntegerNumberConverter;
use crate::converter::integer::integer_string_converter::IntegerStringConverter;
use crate::converter::localdate::local_date_date_converter::LocalDateDateConverter;
use crate::converter::localdate::local_date_number_converter::LocalDateNumberConverter;
use crate::converter::localdate::local_date_string_converter::LocalDateStringConverter;
use crate::converter::localdatetime::local_date_time_date_converter::LocalDateTimeDateConverter;
use crate::converter::localdatetime::local_date_time_number_converter::LocalDateTimeNumberConverter;
use crate::converter::localdatetime::local_date_time_string_converter::LocalDateTimeStringConverter;
use crate::converter::longconverter::long_boolean_converter::LongBooleanConverter;
use crate::converter::longconverter::long_number_converter::LongNumberConverter;
use crate::converter::longconverter::long_string_converter::LongStringConverter;
use crate::converter::shortconverter::short_boolean_converter::ShortBooleanConverter;
use crate::converter::shortconverter::short_number_converter::ShortNumberConverter;
use crate::converter::shortconverter::short_string_converter::ShortStringConverter;
use crate::converter::string::string_boolean_converter::StringBooleanConverter;
use crate::converter::string::string_error_converter::StringErrorConverter;
use crate::converter::string::string_number_converter::StringNumberConverter;
use crate::converter::string::string_string_converter::StringStringConverter;
use crate::converter_registry::ConverterRegistry;
use crate::{
    CellDataType, FromExcelCell, ImageInputStream, IntoExcelCell, JavaDate, ReadConverterContext,
    WriteCellData, WriteConverterContext,
};

struct DefaultReadConverter<T> {
    cell_type: CellDataType,
    marker: PhantomData<fn() -> T>,
}

impl<T> DefaultReadConverter<T> {
    const fn new(cell_type: CellDataType) -> Self {
        Self {
            cell_type,
            marker: PhantomData,
        }
    }
}

impl<T> Converter<T> for DefaultReadConverter<T>
where
    T: FromExcelCell,
{
    fn support_excel_type(&self) -> CellDataType {
        self.cell_type
    }

    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> crate::Result<T> {
        T::from_excel_cell(context.cell(), context.convert_context())
    }
}

struct DefaultWriteConverter<T>(PhantomData<fn() -> T>);

struct DefaultWriteStringConverter<T>(PhantomData<fn() -> T>);

impl<T> Default for DefaultWriteConverter<T> {
    fn default() -> Self {
        Self(PhantomData)
    }
}

impl<T> Default for DefaultWriteStringConverter<T> {
    fn default() -> Self {
        Self(PhantomData)
    }
}

impl<T> Converter<T> for DefaultWriteConverter<T>
where
    T: IntoExcelCell,
{
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, T>,
    ) -> crate::Result<WriteCellData> {
        context
            .value()
            .to_excel_cell(context.convert_context())
            .map(WriteCellData::new)
    }
}

impl<T> Converter<T> for DefaultWriteStringConverter<T>
where
    T: IntoExcelCell,
{
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, T>,
    ) -> crate::Result<WriteCellData> {
        let value = context.value().to_excel_cell(context.convert_context())?;
        let text = match &value {
            crate::CellValue::Date(value) => context.convert_context().format.map_or_else(
                || value.format("%Y-%m-%d").to_string(),
                |format| value.format(format).to_string(),
            ),
            crate::CellValue::DateTime(value) => context.convert_context().format.map_or_else(
                || value.format("%Y-%m-%d %H:%M:%S").to_string(),
                |format| value.format(format).to_string(),
            ),
            _ => value.as_text(),
        };
        Ok(WriteCellData::new(crate::CellValue::String(text)))
    }
}

/// Mirrors Java `DefaultConverterLoader.loadDefaultWriteConverter()`.
pub fn load_default_write_converter() -> ConverterRegistry {
    let mut registry = ConverterRegistry::default();
    macro_rules! register {
        ($($target:ty),+ $(,)?) => {
            $(
                registry.register::<$target, _>(DefaultWriteConverter::<$target>::default());
            )+
        };
    }
    // Java registers one unqualified default write converter per Java class.
    // Rust adds the unsigned/platform integer equivalents because they have no
    // Java primitive counterpart but share the same `IntoExcelCell` contract.
    register!(
        BigDecimal,
        BigInt,
        bool,
        i8,
        i16,
        i32,
        i64,
        isize,
        u8,
        u16,
        u32,
        u64,
        usize,
        f32,
        f64,
        String,
        PathBuf,
        Vec<u8>,
        Box<[u8]>,
        Url,
    );
    registry.register::<bool, _>(BooleanBooleanConverter);
    registry.register::<BigDecimal, _>(BigDecimalNumberConverter);
    registry.register::<BigInt, _>(BigIntegerNumberConverter);
    registry.register::<i8, _>(ByteNumberConverter);
    registry.register::<i16, _>(ShortNumberConverter);
    registry.register::<i32, _>(IntegerNumberConverter);
    registry.register::<i64, _>(LongNumberConverter);
    registry.register::<f32, _>(FloatNumberConverter);
    registry.register::<f64, _>(DoubleNumberConverter);
    registry.register::<JavaDate, _>(DateDateConverter);
    registry.register::<NaiveDate, _>(LocalDateDateConverter);
    registry.register::<NaiveDateTime, _>(LocalDateTimeDateConverter);
    macro_rules! register_string {
        ($($target:ty),+ $(,)?) => {
            $(
                registry.register_for_write_type::<$target, _>(
                    CellDataType::String,
                    DefaultWriteStringConverter::<$target>::default()
                );
            )+
        };
    }
    register_string!(
        BigDecimal, BigInt, bool, i8, i16, i32, i64, isize, u8, u16, u32, u64, usize, f32, f64,
        String
    );
    registry.register::<String, _>(StringStringConverter);
    registry.register::<Vec<u8>, _>(ByteArrayImageConverter);
    registry.register::<Box<[u8]>, _>(BoxingByteArrayImageConverter);
    registry.register::<PathBuf, _>(FileImageConverter);
    registry.register::<ImageInputStream, _>(InputStreamImageConverter);
    registry.register::<Url, _>(crate::UrlImageConverter::default());
    registry.register_for_write_type::<String, _>(CellDataType::String, StringStringConverter);
    registry.register_for_write_type::<JavaDate, _>(CellDataType::String, DateStringConverter);
    registry
        .register_for_write_type::<NaiveDate, _>(CellDataType::String, LocalDateStringConverter);
    registry.register_for_write_type::<NaiveDateTime, _>(
        CellDataType::String,
        LocalDateTimeStringConverter,
    );
    registry.register_for_write_type::<bool, _>(CellDataType::String, BooleanStringConverter);
    registry
        .register_for_write_type::<BigDecimal, _>(CellDataType::String, BigDecimalStringConverter);
    registry.register_for_write_type::<BigInt, _>(CellDataType::String, BigIntegerStringConverter);
    registry.register_for_write_type::<i8, _>(CellDataType::String, ByteStringConverter);
    registry.register_for_write_type::<i16, _>(CellDataType::String, ShortStringConverter);
    registry.register_for_write_type::<i32, _>(CellDataType::String, IntegerStringConverter);
    registry.register_for_write_type::<i64, _>(CellDataType::String, LongStringConverter);
    registry.register_for_write_type::<f32, _>(CellDataType::String, FloatStringConverter);
    registry.register_for_write_type::<f64, _>(CellDataType::String, DoubleStringConverter);
    registry
}

/// Mirrors Java `DefaultConverterLoader.loadDefaultReadConverter()`.
pub fn load_default_read_converter() -> ConverterRegistry {
    let mut registry = ConverterRegistry::default();
    macro_rules! register {
        ($target:ty => [$($cell_type:expr),+ $(,)?]) => {
            $(
                registry.register::<$target, _>(
                    DefaultReadConverter::<$target>::new($cell_type)
                );
            )+
        };
    }
    macro_rules! register_numeric {
        ($($target:ty),+ $(,)?) => {
            $(
                register!($target => [
                    CellDataType::Boolean,
                    CellDataType::Number,
                    CellDataType::String
                ]);
            )+
        };
    }

    register_numeric!(
        BigDecimal, BigInt, i8, i16, i32, i64, isize, u8, u16, u32, u64, usize, f32, f64
    );
    registry.register::<BigDecimal, _>(BigDecimalBooleanConverter);
    registry.register::<BigInt, _>(BigIntegerBooleanConverter);
    registry.register::<i8, _>(ByteBooleanConverter);
    registry.register::<i16, _>(ShortBooleanConverter);
    registry.register::<i32, _>(IntegerBooleanConverter);
    registry.register::<i64, _>(LongBooleanConverter);
    registry.register::<f32, _>(FloatBooleanConverter);
    registry.register::<f64, _>(DoubleBooleanConverter);
    registry.register::<BigDecimal, _>(BigDecimalNumberConverter);
    registry.register::<BigInt, _>(BigIntegerNumberConverter);
    registry.register::<i8, _>(ByteNumberConverter);
    registry.register::<i16, _>(ShortNumberConverter);
    registry.register::<i32, _>(IntegerNumberConverter);
    registry.register::<i64, _>(LongNumberConverter);
    registry.register::<f32, _>(FloatNumberConverter);
    registry.register::<f64, _>(DoubleNumberConverter);
    registry.register::<BigDecimal, _>(BigDecimalStringConverter);
    registry.register::<BigInt, _>(BigIntegerStringConverter);
    registry.register::<i8, _>(ByteStringConverter);
    registry.register::<i16, _>(ShortStringConverter);
    registry.register::<i32, _>(IntegerStringConverter);
    registry.register::<i64, _>(LongStringConverter);
    registry.register::<f32, _>(FloatStringConverter);
    registry.register::<f64, _>(DoubleStringConverter);
    register!(bool => [
        CellDataType::Boolean,
        CellDataType::Number,
        CellDataType::String
    ]);
    registry.register::<bool, _>(BooleanBooleanConverter);
    registry.register::<bool, _>(BooleanNumberConverter);
    registry.register::<bool, _>(BooleanStringConverter);
    registry.register::<JavaDate, _>(DateNumberConverter);
    registry.register::<JavaDate, _>(DateStringConverter);
    register!(NaiveDate => [CellDataType::Date]);
    registry.register::<NaiveDate, _>(LocalDateNumberConverter);
    registry.register::<NaiveDate, _>(LocalDateStringConverter);
    register!(NaiveDateTime => [CellDataType::Date]);
    registry.register::<NaiveDateTime, _>(LocalDateTimeNumberConverter);
    registry.register::<NaiveDateTime, _>(LocalDateTimeStringConverter);
    register!(String => [
        CellDataType::Boolean,
        CellDataType::Number,
        CellDataType::String,
        CellDataType::Error,
        CellDataType::Date,
        CellDataType::RichTextString
    ]);
    registry.register::<String, _>(StringBooleanConverter);
    registry.register::<String, _>(StringNumberConverter);
    registry.register::<String, _>(StringStringConverter);
    registry.register::<String, _>(StringErrorConverter);
    registry
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellValue, ConvertContext, ExcelColumn};
    use std::io::{Cursor, Read};

    fn context() -> ConvertContext {
        ConvertContext {
            sheet_name: "Data".to_owned(),
            row_index: 1,
            column_index: Some(0),
            field: "value",
            format: None,
            use_1904_windowing: false,
        }
    }

    #[test]
    fn default_write_registry_contains_real_scalar_converters() {
        let registry = load_default_write_converter();
        assert!(!registry.is_empty());
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        assert_eq!(
            registry
                .convert_to_excel_data(&42_i32, &column, &context())
                .expect("default i32 converter")
                .expect("i32 converter registered")
                .value(),
            &CellValue::Decimal(BigDecimal::from(42))
        );
        assert_eq!(
            registry
                .convert_to_excel_data(&"text".to_owned(), &column, &context())
                .expect("default String converter")
                .expect("String converter registered")
                .value(),
            &CellValue::String("text".to_owned())
        );
        assert_eq!(
            registry
                .with_write_target(Some(CellDataType::String))
                .convert_to_excel_data(&42_i32, &column, &context())
                .expect("target String conversion")
                .expect("target String converter")
                .value(),
            &CellValue::String("42".to_owned())
        );
        assert_eq!(
            registry
                .with_write_target(Some(CellDataType::String))
                .convert_to_excel_data(&1.0_f64, &column, &context())
                .expect("Java DoubleStringConverter")
                .expect("Double String converter registered")
                .value(),
            &CellValue::String("1.0".to_owned())
        );
        let percent = ExcelColumn::new("value", "Value", Some(0), 0, Some("#.##%"));
        assert_eq!(
            registry
                .with_write_target(Some(CellDataType::String))
                .convert_to_excel_data(&1.235_f64, &percent, &context())
                .expect("formatted Java DoubleStringConverter")
                .expect("formatted Double String converter registered")
                .value(),
            &CellValue::String("123.5%".to_owned())
        );

        let datetime = NaiveDate::from_ymd_opt(2026, 7, 24)
            .unwrap()
            .and_hms_opt(12, 34, 56)
            .unwrap();
        let java_date = JavaDate::from(datetime);
        assert_eq!(
            registry
                .convert_to_excel_data(&java_date, &column, &context())
                .expect("default java.util.Date converter")
                .expect("JavaDate converter registered")
                .value(),
            &CellValue::DateTime(datetime)
        );
        assert_eq!(
            registry
                .convert_to_excel_data(&datetime, &column, &context())
                .expect("default LocalDateTime converter")
                .expect("NaiveDateTime converter registered")
                .value(),
            &CellValue::DateTime(datetime)
        );

        let bytes = vec![0x89, b'P', b'N', b'G'];
        let image = registry
            .convert_to_excel_data(&bytes, &column, &context())
            .expect("Java ByteArrayImageConverter")
            .expect("byte[] converter registered");
        assert_eq!(image.value(), &CellValue::Empty);
        assert_eq!(image.images()[0].image(), bytes);

        let boxed = bytes.clone().into_boxed_slice();
        let image = registry
            .convert_to_excel_data(&boxed, &column, &context())
            .expect("Java BoxingByteArrayImageConverter")
            .expect("Byte[] converter registered");
        assert_eq!(image.images()[0].image(), bytes);

        let path = std::env::temp_dir().join(format!(
            "easyexcel-file-image-converter-{}",
            std::process::id()
        ));
        std::fs::write(&path, &bytes).expect("write image fixture");
        let image = registry
            .convert_to_excel_data(&path, &column, &context())
            .expect("Java FileImageConverter")
            .expect("File converter registered");
        std::fs::remove_file(path).expect("remove image fixture");
        assert_eq!(image.images()[0].image(), bytes);

        let stream: ImageInputStream = ImageInputStream::boxed(Cursor::new(bytes.clone()));
        let image = registry
            .convert_to_excel_data(&stream, &column, &context())
            .expect("Java InputStreamImageConverter")
            .expect("InputStream converter registered");
        assert_eq!(image.images()[0].image(), bytes);
        let mut inner = stream.into_inner();
        let mut remaining = Vec::new();
        inner
            .read_to_end(&mut remaining)
            .expect("inspect consumed stream");
        assert!(
            remaining.is_empty(),
            "conversion consumes the current stream remainder exactly once"
        );
    }

    #[test]
    fn default_read_registry_dispatches_by_target_and_cell_type() {
        let registry = load_default_read_converter();
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        let convert_context = context();
        let boolean = CellValue::Bool(true);
        let read_context = ReadConverterContext::new(Some(&boolean), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<i32>(&read_context)
                .expect("default i32 converter"),
            Some(1)
        );
        assert_eq!(
            registry
                .convert_to_rust_data::<String>(&read_context)
                .expect("default String converter"),
            Some("true".to_owned())
        );
        let error = CellValue::Error("#DIV/0!".to_owned());
        let read_context = ReadConverterContext::new(Some(&error), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<String>(&read_context)
                .expect("Java StringErrorConverter"),
            Some("#DIV/0!".to_owned())
        );
        let text = CellValue::String("exact".to_owned());
        let read_context = ReadConverterContext::new(Some(&text), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<String>(&read_context)
                .expect("Java StringStringConverter"),
            Some("exact".to_owned())
        );
        let number = CellValue::Float(1.5);
        let read_context = ReadConverterContext::new(Some(&number), &column, &convert_context);
        let expected = NaiveDate::from_ymd_opt(1900, 1, 1)
            .unwrap()
            .and_hms_opt(12, 0, 0)
            .unwrap();
        assert_eq!(
            registry
                .convert_to_rust_data::<JavaDate>(&read_context)
                .expect("default java.util.Date converter")
                .expect("JavaDate number converter registered")
                .naive_local(),
            expected
        );
        assert_eq!(
            registry
                .convert_to_rust_data::<NaiveDateTime>(&read_context)
                .expect("default LocalDateTime converter"),
            Some(expected)
        );
        let two = CellValue::Int(2);
        let read_context = ReadConverterContext::new(Some(&two), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<bool>(&read_context)
                .expect("Java BooleanNumberConverter"),
            Some(false)
        );
        let one_text = CellValue::String("1".to_owned());
        let read_context = ReadConverterContext::new(Some(&one_text), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<bool>(&read_context)
                .expect("Java BooleanStringConverter"),
            Some(false)
        );
        let wrapped = CellValue::Decimal(
            "4294967295.9"
                .parse()
                .expect("valid Java BigDecimal fixture"),
        );
        let read_context = ReadConverterContext::new(Some(&wrapped), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<i32>(&read_context)
                .expect("Java IntegerNumberConverter"),
            Some(-1)
        );
        let wrapped_text = CellValue::String("4294967295.9".to_owned());
        let read_context =
            ReadConverterContext::new(Some(&wrapped_text), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<i32>(&read_context)
                .expect("Java IntegerStringConverter"),
            Some(-1)
        );

        let serial = CellValue::Float(1.5);
        let date_context = ReadConverterContext::new(Some(&serial), &column, &convert_context);
        assert_eq!(
            registry
                .convert_to_rust_data::<NaiveDateTime>(&date_context)
                .expect("1900 datetime converter"),
            Some(
                NaiveDate::from_ymd_opt(1900, 1, 1)
                    .expect("date")
                    .and_hms_opt(12, 0, 0)
                    .expect("time")
            )
        );

        let mut window_1904 = convert_context.clone();
        window_1904.use_1904_windowing = true;
        let zero = CellValue::Int(0);
        let date_context = ReadConverterContext::new(Some(&zero), &column, &window_1904);
        assert_eq!(
            registry
                .convert_to_rust_data::<NaiveDate>(&date_context)
                .expect("1904 date converter"),
            NaiveDate::from_ymd_opt(1904, 1, 1)
        );
    }
}
