//! Mirrors Java `com.alibaba.excel.event.Listener`.
//!
//! Java `Listener` is an empty marker interface. Rust mirrors it as an
/// empty trait so 1:1 Java package references resolve.

/// Empty marker trait mirroring Java `Listener`.
pub trait Listener {}
