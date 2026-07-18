//! Mirrors Java com.alibaba.excel.util.Validate.

#![allow(dead_code)]

use crate::excel_error::ExcelError;

/// Mirrors `org.apache.commons.lang3.Validate#isTrue`.
pub fn is_true(expression: bool, message: impl Into<String>) -> Result<(), ExcelError> {
    if expression {
        Ok(())
    } else {
        Err(ExcelError::Unsupported(message.into()))
    }
}
