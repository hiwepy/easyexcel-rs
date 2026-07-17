//! Mirrors Java `com.alibaba.excel.enums.CellDataTypeEnum`.
//!
//! Java 定义了 8 个变体；Rust 额外补齐了 `Formula` 和 `Image`，与 `CellValue`
//! 中 `Formula(String) / Image(Vec<u8>)` 变体对齐。
//!
//! 原 Java `buildFromCellType(String)` 通过类型码 `"s" / "str" / "inlineStr" / "e" / "b" / "n"`
//! 路由到枚举；这一逻辑在 `easyexcel-reader/src/xlsx_rows.rs::finish_cell` 内联。

/// Logical Excel cell type used as the read-converter dispatch key.
///
/// This is the Rust port of Java `CellDataTypeEnum`. Two additional variants
/// (`Formula`, `Image`) keep it aligned with the `CellValue` enum so writers
/// can carry Java-equivalent rich metadata without an extra wrapper class.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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
