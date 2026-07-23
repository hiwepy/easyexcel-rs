//! Holder 接口镜像。
//!
//! 对应 Java：`com.alibaba.excel.metadata.Holder`
//! Java 枚举 `HolderEnum` 在 Rust 中实现为 [`crate::Holder`]（`enum_holder.rs`）。

use crate::Holder;

/// Java `Holder` 接口的 Rust trait。
///
/// # Java 对应
/// - 接口：`com.alibaba.excel.metadata.Holder`
/// - 方法：`HolderEnum holderType()` → [`Self::holder_type`]
pub trait ExcelHolder {
    /// 返回 holder 类型。对应 Java `holderType()`。
    fn holder_type(&self) -> Holder;
}

/// Java `HolderEnum` 命名别名（Rust 类型为 [`Holder`]）。
pub type HolderEnum = Holder;
