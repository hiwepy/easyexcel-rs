//! Mirrors Java `com.alibaba.excel.event.Handler`.
//!
//! Java `Handler extends Order`. The `order()` method has a default of
//! `OrderConstant.DEFAULT_ORDER` (= 0). Rust already encodes this on
//! `WriteHandler::order()`. This module re-exports the same value as a
//! standalone trait so 1:1 Java package references resolve.

/// Mirrors Java `Handler extends Order`.
///
/// `Handler` is a marker extension of `Order`; Rust mirrors the
/// contract through the `order()` method.
pub trait Handler {
    /// Returns the handler's execution order. Lower values execute first.
    /// (Java `Handler.order()` defaulting to `OrderConstant.DEFAULT_ORDER`)
    fn order(&self) -> i32 {
        0
    }
}
