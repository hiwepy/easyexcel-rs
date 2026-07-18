//! Mirrors Java com.alibaba.excel.util.NumberUtils.

#![allow(dead_code)]

use bigdecimal::BigDecimal;
use num_bigint::BigInt;

use crate::excel_error::ExcelError;

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#format`.
#[must_use]
pub fn format(value: f64, scale: usize) -> String {
    format!("{value:.*}", scale)
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseShort`.
pub fn parse_short(str: &str) -> Result<i16, ExcelError> {
    str.trim()
        .parse::<i16>()
        .map_err(|_| ExcelError::Format(format!("parseShort failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseLong`.
pub fn parse_long(str: &str) -> Result<i64, ExcelError> {
    str.trim()
        .parse::<i64>()
        .map_err(|_| ExcelError::Format(format!("parseLong failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseInteger`.
pub fn parse_integer(str: &str) -> Result<i32, ExcelError> {
    str.trim()
        .parse::<i32>()
        .map_err(|_| ExcelError::Format(format!("parseInteger failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseFloat`.
pub fn parse_float(str: &str) -> Result<f32, ExcelError> {
    str.trim()
        .parse::<f32>()
        .map_err(|_| ExcelError::Format(format!("parseFloat failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseBigDecimal`.
pub fn parse_big_decimal(str: &str) -> Result<BigDecimal, ExcelError> {
    str.trim()
        .parse::<BigDecimal>()
        .map_err(|_| ExcelError::Format(format!("parseBigDecimal failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseByte`.
pub fn parse_byte(str: &str) -> Result<i8, ExcelError> {
    str.trim()
        .parse::<i8>()
        .map_err(|_| ExcelError::Format(format!("parseByte failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#parseDouble`.
pub fn parse_double(str: &str) -> Result<f64, ExcelError> {
    str.trim()
        .parse::<f64>()
        .map_err(|_| ExcelError::Format(format!("parseDouble failed for {str:?}")))
}

/// Mirrors `org.apache.commons.lang3.math.NumberUtils#createBigInteger`.
pub fn parse_big_int(str: &str) -> Result<BigInt, ExcelError> {
    str.trim()
        .parse::<BigInt>()
        .map_err(|_| ExcelError::Format(format!("parseBigInteger failed for {str:?}")))
}
