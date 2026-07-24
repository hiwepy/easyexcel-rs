//! Mirrors the union of `com.alibaba.excel.converters.*.java` (the ~40
//! built-in `Converter<T>` implementations registered by Java's
//! `DefaultConverterLoader`).
//!
//! Each `impl FromExcelCell for X` and `impl IntoExcelCell for X` here
//! corresponds to a Java converter under
//! `com.alibaba.excel.converters.{bigdecimal,biginteger,booleanconverter,
//!  byteconverter,date,doubleconverter,floatconverter,integer,localdate,
//!  localdatetime,longconverter,shortconverter,string}` plus the
//! `Vec<u8>` / `Box<[u8]>` / `[u8; N]` / `PathBuf` image converters.

use std::fmt::Display;
use std::str::FromStr;

use bigdecimal::BigDecimal;
use bigdecimal::ToPrimitive;
use chrono::{Duration, NaiveDate, NaiveDateTime};
use num_bigint::BigInt;

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::dynamic_row::DynamicRow;
use crate::excel_error::ExcelError;
use crate::excel_row::ExcelRow;
use crate::from_excel_cell::FromExcelCell;
use crate::into_excel_cell::IntoExcelCell;
use crate::row_data::RowData;

impl FromExcelCell for String {
    fn from_excel_cell(
        value: Option<&CellValue>,
        _context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        Ok(value.map_or_else(String::new, CellValue::as_text))
    }
}

impl IntoExcelCell for String {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::String(self.clone()))
    }
}

impl IntoExcelCell for &str {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::String((*self).to_owned()))
    }
}

impl FromExcelCell for bool {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        match value.unwrap_or(&CellValue::Empty) {
            CellValue::Bool(value) => Ok(*value),
            CellValue::Int(value) => Ok(*value != 0),
            CellValue::Float(value) => Ok(*value != 0.0),
            CellValue::Decimal(value) => Ok(value != &BigDecimal::from(0)),
            CellValue::String(value) if value.eq_ignore_ascii_case("true") || value == "1" => {
                Ok(true)
            }
            CellValue::String(value) if value.eq_ignore_ascii_case("false") || value == "0" => {
                Ok(false)
            }
            other => Err(context.invalid(other, "bool")),
        }
    }
}

impl IntoExcelCell for bool {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Bool(*self))
    }
}

macro_rules! integer_conversion {
    ($($ty:ty),+ $(,)?) => {
        $(
            impl FromExcelCell for $ty {
                fn from_excel_cell(
                    value: Option<&CellValue>,
                    context: &ConvertContext,
                ) -> Result<Self, ExcelError> {
                    parse_integer(value, context, stringify!($ty))
                }
            }

            impl IntoExcelCell for $ty {
                fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
                    Ok(integer_to_cell(*self))
                }
            }
        )+
    };
}

integer_conversion!(i8, i16, i32, i64, isize, u8, u16, u32, u64, usize);

impl FromExcelCell for BigInt {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        let cell = value.unwrap_or(&CellValue::Empty);
        match cell {
            CellValue::Bool(value) => Ok(Self::from(u8::from(*value))),
            CellValue::Int(value) => Ok(Self::from(*value)),
            CellValue::Float(value) => BigDecimal::from_str(&value.to_string())
                .map(|value| decimal_to_big_int(&value))
                .map_err(|_| context.invalid(cell, "BigInt")),
            CellValue::Decimal(value) => Ok(decimal_to_big_int(value)),
            CellValue::String(value) => BigDecimal::from_str(value)
                .map(|value| decimal_to_big_int(&value))
                .map_err(|_| context.invalid(cell, "BigInt")),
            other => Err(context.invalid(other, "BigInt")),
        }
    }
}

impl IntoExcelCell for BigInt {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(self
            .to_i64()
            .map_or_else(|| CellValue::String(self.to_string()), CellValue::Int))
    }
}

fn decimal_to_big_int(value: &BigDecimal) -> BigInt {
    value.with_scale(0).into_bigint_and_exponent().0
}

fn parse_integer<T>(
    value: Option<&CellValue>,
    context: &ConvertContext,
    target: &'static str,
) -> Result<T, ExcelError>
where
    T: FromStr,
{
    let value = value.unwrap_or(&CellValue::Empty);
    let text = match value {
        CellValue::Bool(inner) => u8::from(*inner).to_string(),
        CellValue::Int(inner) => inner.to_string(),
        CellValue::Float(inner) if inner.fract() == 0.0 => inner.to_string(),
        CellValue::Decimal(inner) if inner == &inner.with_scale(0) => inner.to_string(),
        CellValue::String(inner) => inner.clone(),
        other => return Err(context.invalid(other, target)),
    };
    text.parse::<T>()
        .map_err(|_| context.invalid(value, target))
}

fn integer_to_cell<T>(value: T) -> CellValue
where
    T: TryInto<i64> + Display + Copy,
{
    value
        .try_into()
        .map_or_else(|_| CellValue::String(value.to_string()), CellValue::Int)
}

macro_rules! float_conversion {
    ($($ty:ty),+ $(,)?) => {
        $(
            impl FromExcelCell for $ty {
                fn from_excel_cell(
                    value: Option<&CellValue>,
                    context: &ConvertContext,
                ) -> Result<Self, ExcelError> {
                    parse_float(value, context, stringify!($ty))
                }
            }

            impl IntoExcelCell for $ty {
                fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
                    Ok(CellValue::Float(f64::from(*self)))
                }
            }
        )+
    };
}

float_conversion!(f32, f64);

fn parse_float<T>(
    value: Option<&CellValue>,
    context: &ConvertContext,
    target: &'static str,
) -> Result<T, ExcelError>
where
    T: FromStr,
{
    let value = value.unwrap_or(&CellValue::Empty);
    let text = match value {
        CellValue::Bool(inner) => u8::from(*inner).to_string(),
        CellValue::Int(inner) => inner.to_string(),
        CellValue::Float(inner) => inner.to_string(),
        CellValue::Decimal(inner) => inner.to_string(),
        CellValue::String(inner) => inner.clone(),
        other => return Err(context.invalid(other, target)),
    };
    text.parse::<T>()
        .map_err(|_| context.invalid(value, target))
}

impl FromExcelCell for BigDecimal {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::Bool(inner) => Ok(Self::from(u8::from(*inner))),
            CellValue::Decimal(inner) => Ok(inner.clone()),
            CellValue::Int(inner) => Ok(Self::from(*inner)),
            CellValue::Float(inner) => {
                Self::from_str(&inner.to_string()).map_err(|_| context.invalid(value, "BigDecimal"))
            }
            CellValue::String(inner) => {
                Self::from_str(inner).map_err(|_| context.invalid(value, "BigDecimal"))
            }
            other => Err(context.invalid(other, "BigDecimal")),
        }
    }
}

impl IntoExcelCell for BigDecimal {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Decimal(self.clone()))
    }
}

impl FromExcelCell for NaiveDate {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::Date(value) => Ok(*value),
            CellValue::DateTime(value) => Ok(value.date()),
            CellValue::Int(_) | CellValue::Float(_) | CellValue::Decimal(_) => {
                excel_serial_to_datetime(value, context).map(|value| value.date())
            }
            CellValue::String(inner) => {
                NaiveDate::parse_from_str(inner, context.format.unwrap_or("%Y-%m-%d"))
                    .map_err(|_| context.invalid(value, "NaiveDate"))
            }
            other => Err(context.invalid(other, "NaiveDate")),
        }
    }
}

impl IntoExcelCell for NaiveDate {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Date(*self))
    }
}

impl FromExcelCell for NaiveDateTime {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::DateTime(value) => Ok(*value),
            CellValue::Date(value) => Ok(value.and_hms_opt(0, 0, 0).expect("midnight is valid")),
            CellValue::Int(_) | CellValue::Float(_) | CellValue::Decimal(_) => {
                excel_serial_to_datetime(value, context)
            }
            CellValue::String(inner) => {
                NaiveDateTime::parse_from_str(inner, context.format.unwrap_or("%Y-%m-%d %H:%M:%S"))
                    .map_err(|_| context.invalid(value, "NaiveDateTime"))
            }
            other => Err(context.invalid(other, "NaiveDateTime")),
        }
    }
}

fn excel_serial_to_datetime(
    value: &CellValue,
    context: &ConvertContext,
) -> Result<NaiveDateTime, ExcelError> {
    let serial = match value {
        CellValue::Int(value) => *value as f64,
        CellValue::Float(value) => *value,
        CellValue::Decimal(decimal) => decimal
            .to_f64()
            .ok_or_else(|| context.invalid(value, "Excel date"))?,
        other => return Err(context.invalid(other, "Excel date")),
    };
    if !serial.is_finite() || serial < 0.0 {
        return Err(context.invalid(value, "Excel date"));
    }

    let whole_days = serial.floor() as i64;
    // POI preserves Excel's fictitious 1900-02-29 by mapping serials 60 and
    // 61 to 1900-03-01. Before 61 the effective epoch is 1899-12-31; from
    // 61 onward it is 1899-12-30.
    let epoch = if context.use_1904_windowing {
        NaiveDate::from_ymd_opt(1904, 1, 1).expect("valid Excel epoch")
    } else if whole_days < 61 {
        NaiveDate::from_ymd_opt(1899, 12, 31).expect("valid Excel epoch")
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).expect("valid Excel epoch")
    };
    let milliseconds = ((serial - serial.floor()) * 86_400_000.0 + 0.5).floor() as i64;
    epoch
        .and_hms_opt(0, 0, 0)
        .and_then(|value| value.checked_add_signed(Duration::days(whole_days)))
        .and_then(|value| value.checked_add_signed(Duration::milliseconds(milliseconds)))
        .ok_or_else(|| context.invalid(value, "Excel date"))
}

impl IntoExcelCell for NaiveDateTime {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::DateTime(*self))
    }
}

impl FromExcelCell for Vec<u8> {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        let value = value.unwrap_or(&CellValue::Empty);
        match value {
            CellValue::Image(bytes) => Ok(bytes.clone()),
            other => Err(context.invalid(other, "Vec<u8>")),
        }
    }
}

impl IntoExcelCell for Vec<u8> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Image(self.clone()))
    }
}

impl FromExcelCell for Box<[u8]> {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        Vec::<u8>::from_excel_cell(value, context).map(Vec::into_boxed_slice)
    }
}

impl IntoExcelCell for Box<[u8]> {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Image(self.to_vec()))
    }
}

impl<const N: usize> FromExcelCell for [u8; N] {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        Vec::<u8>::from_excel_cell(value, context)?
            .try_into()
            .map_err(|_| context.invalid(value.unwrap_or(&CellValue::Empty), "[u8; N]"))
    }
}

impl<const N: usize> IntoExcelCell for [u8; N] {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        Ok(CellValue::Image(self.to_vec()))
    }
}

impl FromExcelCell for std::path::PathBuf {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        String::from_excel_cell(value, context).map(Self::from)
    }
}

impl IntoExcelCell for std::path::PathBuf {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        std::fs::read(self)
            .map(CellValue::Image)
            .map_err(Into::into)
    }
}

impl<T: FromExcelCell> FromExcelCell for Option<T> {
    fn from_excel_cell(
        value: Option<&CellValue>,
        context: &ConvertContext,
    ) -> Result<Self, ExcelError> {
        if value.is_none_or(CellValue::is_empty) {
            Ok(None)
        } else {
            T::from_excel_cell(value, context).map(Some)
        }
    }
}

impl<T: IntoExcelCell> IntoExcelCell for Option<T> {
    fn to_excel_cell(&self, context: &ConvertContext) -> Result<CellValue, ExcelError> {
        self.as_ref()
            .map_or(Ok(CellValue::Empty), |value| value.to_excel_cell(context))
    }
}

impl ExcelRow for DynamicRow {
    fn schema() -> &'static [crate::excel_column::ExcelColumn] {
        &[]
    }

    fn from_row(row: &RowData) -> Result<Self, ExcelError> {
        Ok(Self(
            (0..row.dynamic_width())
                .map(|index| (index, row.dynamic_cell(index)))
                .collect(),
        ))
    }

    fn to_row(&self) -> Result<Vec<CellValue>, ExcelError> {
        let entries = &self.0;
        let Some(last_index) = entries.last_key_value().map(|(index, _)| *index) else {
            return Ok(Vec::new());
        };
        let row_length = last_index
            .checked_add(1)
            .ok_or_else(|| ExcelError::Format("dynamic column index exceeds usize".to_owned()))?;
        let mut row = vec![CellValue::Empty; row_length];
        for (index, value) in entries {
            row[*index] = match value {
                crate::dynamic_value::DynamicValue::Null => CellValue::Empty,
                crate::dynamic_value::DynamicValue::String(value) => {
                    CellValue::String(value.clone())
                }
                crate::dynamic_value::DynamicValue::ActualData(value) => value.clone(),
                crate::dynamic_value::DynamicValue::ReadCellData(value) => value.data().clone(),
            };
        }
        Ok(row)
    }
}
