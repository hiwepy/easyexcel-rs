//! Mirrors Java `com.alibaba.excel.event.Order`.

/// Implement this interface when sorting.
///
/// Rust port of Java `Order`. The smaller the first implementation.
pub trait Order {
    /// Returns the sort order. Lower values execute first.
    fn order(&self) -> i32;
}
