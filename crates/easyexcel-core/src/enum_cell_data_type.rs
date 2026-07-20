//! Mirrors Java `com.alibaba.excel.enums.CellDataTypeEnum`.
//!
//! Java 定义了 8 个变体；Rust 额外补齐了 `Formula` 和 `Image`，与 `CellValue`
//! 中 `Formula(String) / Image(Vec<u8>)` 变体对齐。
//!
//! 原 Java `buildFromCellType(String)` 通过类型码 `"s" / "str" / "inlineStr" / "e" / "b" / "n"`
//! 路由到枚举；见 [`CellDataType::build_from_cell_type`]。

/// Logical Excel cell type used as the read-converter dispatch key.
///
/// This is the Rust port of Java `CellDataTypeEnum`. Two additional variants
/// (`Formula`, `Image`) keep it aligned with the `CellValue` enum so writers
/// can carry Java-equivalent rich metadata without an extra wrapper class.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum CellDataType {
    /// Shared or inline string.       (Java `STRING`)
    String,
    /// Direct inline string.          (Java `DIRECT_STRING`)
    DirectString,
    /// Numeric value.                  (Java `NUMBER`)
    Number,
    /// Boolean value.                  (Java `BOOLEAN`)
    Boolean,
    /// Empty or physically absent cell. (Java `EMPTY`)
    #[default]
    Empty,
    /// Excel error value.              (Java `ERROR`)
    Error,
    /// Date or date-time value.        (Java `DATE`)
    Date,
    /// Rich text string.              (Java `RICH_TEXT_STRING`)
    RichTextString,
    /// Formula expression supplied as a write value. (Rust extension)
    Formula,
    /// Encoded image bytes.            (Rust extension)
    Image,
}

impl CellDataType {
    /// Java `CellDataTypeEnum.buildFromCellType(String)`.
    ///
    /// Maps OOXML `c@t` codes onto the enum used by `CellTagHandler.startElement`.
    /// Unknown codes return [`None`] (Java would return `null` and later NPE —
    /// Rust callers treat that as a format error).
    #[must_use]
    pub fn build_from_cell_type(cell_type: Option<&str>) -> Option<Self> {
        match cell_type {
            None | Some("") => Some(Self::Empty),
            Some("s") => Some(Self::String),
            Some("str" | "inlineStr") => Some(Self::DirectString),
            Some("e") => Some(Self::Error),
            Some("b") => Some(Self::Boolean),
            Some("n") => Some(Self::Number),
            // Rust path also accepts date serials marked `d` (OOXML extension).
            Some("d") => Some(Self::DirectString),
            Some(_) => None,
        }
    }
}
