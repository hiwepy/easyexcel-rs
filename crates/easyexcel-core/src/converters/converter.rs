//! 转换器接口。
//!
//! 对应 Java：`com.alibaba.excel.converters.Converter`
//! 既有实现位于 [`crate::converter::converter_trait`]，本文件做 1:1 路径 re-export。

#![allow(unused_imports)]
pub use crate::converter::converter_trait::Converter;
