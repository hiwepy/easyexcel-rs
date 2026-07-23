//! 字体元数据（已弃用）。
//!
//! 对应 Java：`com.alibaba.excel.metadata.Font`（`@Deprecated`，请改用 `WriteFont`）。
//! 原 Java 文件：`easyexcel-core/.../metadata/Font.java`

use crate::WriteFont;

/// 已弃用的字体模型，对齐 Java `Font`。
///
/// # Java 对应
/// - 类：`com.alibaba.excel.metadata.Font`
/// - 替代：`com.alibaba.excel.write.metadata.style.WriteFont` → [`WriteFont`]
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Font {
    /// 字体名称。Java `fontName` / `getFontName()` / `setFontName`
    font_name: Option<String>,
    /// 字号（磅）。Java `fontHeightInPoints`
    font_height_in_points: i16,
    /// 是否加粗。Java `bold` / `isBold()` / `setBold`
    bold: bool,
}

impl Font {
    /// 创建空字体。对应 Java 默认构造。
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// 返回字体名称。对应 Java `getFontName()`。
    #[must_use]
    pub fn font_name(&self) -> Option<&str> {
        self.font_name.as_deref()
    }

    /// 设置字体名称。对应 Java `setFontName(String)`。
    pub fn set_font_name(&mut self, font_name: impl Into<String>) {
        self.font_name = Some(font_name.into());
    }

    /// 返回字号。对应 Java `getFontHeightInPoints()`。
    #[must_use]
    pub const fn font_height_in_points(&self) -> i16 {
        self.font_height_in_points
    }

    /// 设置字号。对应 Java `setFontHeightInPoints(short)`。
    pub const fn set_font_height_in_points(&mut self, value: i16) {
        self.font_height_in_points = value;
    }

    /// 是否加粗。对应 Java `isBold()`。
    #[must_use]
    pub const fn is_bold(&self) -> bool {
        self.bold
    }

    /// 设置加粗。对应 Java `setBold(boolean)`。
    pub const fn set_bold(&mut self, bold: bool) {
        self.bold = bold;
    }

    /// 转换为推荐的 [`WriteFont`]。
    #[must_use]
    pub fn to_write_font(&self) -> WriteFont {
        let mut font = WriteFont::new();
        if let Some(name) = &self.font_name {
            font = font.font_name(name.clone());
        }
        font = font.font_height_in_points(f64::from(self.font_height_in_points));
        font.bold(self.bold)
    }
}
