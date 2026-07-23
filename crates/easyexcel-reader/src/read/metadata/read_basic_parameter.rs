//! 读取基础参数。
//!
//! 对应 Java：`com.alibaba.excel.read.metadata.ReadBasicParameter`
//! 原文件：`easyexcel-core/.../read/metadata/ReadBasicParameter.java`
//!
//! 继承面：Java extends `BasicParameter` → 组合 [`easyexcel_core::BasicParameter`]。

use easyexcel_core::metadata::BasicParameter;

/// 读取基础参数，对齐 Java `ReadBasicParameter`。
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ReadBasicParameter {
    /// 基础参数（Java 继承字段）。对应 `BasicParameter`
    pub basic: BasicParameter,
    /// 表头行数。Java `headRowNumber` / `getHeadRowNumber()` / `setHeadRowNumber`
    pub head_row_number: Option<i32>,
    /// 自定义监听器类型名列表（Rust 无反射，用类型名占位）。
    /// Java `customReadListenerList`
    pub custom_read_listener_list: Vec<String>,
}

impl ReadBasicParameter {
    /// 创建参数。对应 Java 构造：初始化空 `customReadListenerList`。
    #[must_use]
    pub fn new() -> Self {
        Self {
            basic: BasicParameter::new(),
            head_row_number: None,
            custom_read_listener_list: Vec::new(),
        }
    }

    /// 返回表头行数。对应 Java `getHeadRowNumber()`。
    #[must_use]
    pub const fn head_row_number(&self) -> Option<i32> {
        self.head_row_number
    }

    /// 设置表头行数。对应 Java `setHeadRowNumber(Integer)`。
    pub fn set_head_row_number(&mut self, value: Option<i32>) {
        self.head_row_number = value;
    }

    /// 返回自定义监听器列表。对应 Java `getCustomReadListenerList()`。
    #[must_use]
    pub fn custom_read_listener_list(&self) -> &[String] {
        &self.custom_read_listener_list
    }
}
