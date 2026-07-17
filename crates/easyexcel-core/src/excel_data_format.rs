//! Mirrors Java `com.alibaba.excel.metadata.data.DataFormatData`.

/// A Java built-in number-format index or a custom Excel format string.
///
/// Java `DataFormatData` carries both an `index: Short` and a `format: String`
/// with a static `merge` helper. Rust keeps the union semantics in a single
/// enum and pushes the merge logic onto `WriteCellStyle.merge` (Java-side
/// equivalent) or into the writer-side helpers.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelDataFormat {
    /// A Java EasyExcel / Apache POI built-in format index.
    Builtin(u8),
    /// A custom Excel number-format string.
    Custom(&'static str),
}
