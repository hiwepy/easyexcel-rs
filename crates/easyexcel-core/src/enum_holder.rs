//! Mirrors Java `com.alibaba.excel.enums.HolderEnum`.

/// The types of holder.
///
/// Rust port of Java `HolderEnum`. Used to tag workbook / sheet / table / row
/// containers, although Rust collapses most of these into `ReadOptions` /
/// `WriteOptions` plus private state inside the writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Holder {
    /// Workbook-scoped holder.
    Workbook,
    /// Sheet-scoped holder.
    Sheet,
    /// Table-scoped holder.
    Table,
    /// Row-scoped holder.
    Row,
}
