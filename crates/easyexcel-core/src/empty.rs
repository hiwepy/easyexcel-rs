//! 空占位类型。
//!
//! 对应 Java：`com.alibaba.excel.Empty` / `easyexcel-support` 的 `Empty`
//! 原文件：聚合模块与 support 模块中的 Empty.java

/// 空标记结构，对齐 Java `Empty`（无字段）。
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Empty;

#[cfg(test)]
mod tests {
    use std::any::TypeId;

    use super::Empty;

    #[test]
    fn root_and_support_empty_are_the_same_zero_sized_marker() {
        let marker = Empty;
        assert_eq!(marker, Empty::default());
        assert_eq!(std::mem::size_of::<Empty>(), 0);
        assert_eq!(
            TypeId::of::<Empty>(),
            TypeId::of::<crate::support::empty::Empty>()
        );
    }
}
