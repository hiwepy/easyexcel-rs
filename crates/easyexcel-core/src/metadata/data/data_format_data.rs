//! 单元格数据格式。
//!
//! 对应 Java：`com.alibaba.excel.metadata.data.DataFormatData`
//! 原文件：`easyexcel-core/.../metadata/data/DataFormatData.java`

use std::borrow::Cow;

/// 数据格式元数据，对齐 Java `DataFormatData`。
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DataFormatData {
    /// 格式索引。Java `index` / `getIndex()` / `setIndex`
    pub index: Option<i16>,
    /// 格式串。Java `format` / `getFormat()` / `setFormat`
    pub format: Option<String>,
}

impl DataFormatData {
    /// 创建空格式。对应 Java 默认构造。
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// 返回格式索引。对应 Java `getIndex()`。
    #[must_use]
    pub const fn index(&self) -> Option<i16> {
        self.index
    }

    /// 设置格式索引。对应 Java `setIndex(Short)`。
    pub fn set_index(&mut self, index: Option<i16>) {
        self.index = index;
    }

    /// 返回格式串。对应 Java `getFormat()`。
    #[must_use]
    pub fn format(&self) -> Option<&str> {
        self.format.as_deref()
    }

    /// 设置格式串。对应 Java `setFormat(String)`。
    pub fn set_format(&mut self, format: impl Into<Option<String>>) {
        self.format = format.into();
    }

    /// 将 source 非空字段合并到 target。对应 Java `merge(source, target)`。
    pub fn merge(source: Option<&Self>, target: Option<&mut Self>) {
        let (Some(source), Some(target)) = (source, target) else {
            return;
        };
        if let Some(index) = source.index {
            target.index = Some(index);
        }
        if let Some(format) = source.format.as_ref().filter(|s| !s.trim().is_empty()) {
            target.format = Some(format.clone());
        }
    }

    /// 克隆副本。对应 Java `clone()`。
    #[must_use]
    pub fn clone_data(&self) -> Self {
        self.clone()
    }

    /// 借用格式文本（测试辅助）。
    #[must_use]
    pub fn format_cow(&self) -> Cow<'_, str> {
        match &self.format {
            Some(s) => Cow::Borrowed(s.as_str()),
            None => Cow::Borrowed(""),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn merge_overwrites_non_empty_fields() {
        let source = DataFormatData {
            index: Some(1),
            format: Some("0.00".to_owned()),
        };
        let mut target = DataFormatData::new();
        DataFormatData::merge(Some(&source), Some(&mut target));
        assert_eq!(target.index, Some(1));
        assert_eq!(target.format.as_deref(), Some("0.00"));
    }
}
