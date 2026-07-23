//! EasyExcel 工厂。
//!
//! 对应 Java：`com.alibaba.excel.EasyExcelFactory`
//! Java 中 `EasyExcel extends EasyExcelFactory`；Rust 合并为同一 [`crate::EasyExcel`]。

#![allow(unused_imports)]
pub use crate::EasyExcel as EasyExcelFactory;
