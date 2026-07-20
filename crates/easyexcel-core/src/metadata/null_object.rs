//! Mirrors Java `com.alibaba.excel.metadata.NullObject`.

/// Null object placeholder.
///
/// Java uses this marker type when a converter or handler needs a non-null
/// sentinel without allocating real data. Rust keeps the same zero-sized
/// marker so call sites can mirror Java `new NullObject()` semantics.
///
/// Rust port of Java `NullObject`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct NullObject;

impl NullObject {
    /// Creates a null-object sentinel. (Java default constructor)
    #[must_use]
    pub const fn new() -> Self {
        Self
    }
}
