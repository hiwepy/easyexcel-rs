//! 空占位类型。
//!
//! 对应 Java：`com.alibaba.excel.Empty` / `easyexcel-support` 的 `Empty`
//! 原文件：聚合模块与 support 模块中的 Empty.java

/// 空标记结构，对齐 Java `Empty`（无字段）。
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Empty;
