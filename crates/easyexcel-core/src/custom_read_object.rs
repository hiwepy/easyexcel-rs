//! Mirrors the `customObject` field of Java `ReadWorkbookHolder` plus
//! `AnalysisContext.getCustom()`.

use std::any::Any;
use std::sync::Arc;

/// Type-safe shared value equivalent to Java `EasyExcel`'s reader `customObject`.
///
/// Java `AnalysisContext.getCustom()` returns `Object`. Rust hides the value
/// behind an `Arc<dyn Any>` so listeners can downcast back to the original
/// type with `downcast_ref`.
#[derive(Clone)]
pub struct CustomReadObject(Arc<dyn Any + Send + Sync>);

impl CustomReadObject {
    /// Wraps a value for propagation to every read callback context.
    #[must_use]
    pub fn new<T>(value: T) -> Self
    where
        T: Any + Send + Sync,
    {
        Self(Arc::new(value))
    }

    /// Returns the value when its concrete type matches `T`. Mirrors Java
    /// `(T) AnalysisContext.getCustom()` after an explicit cast.
    #[must_use]
    pub fn downcast_ref<T: Any>(&self) -> Option<&T> {
        self.0.downcast_ref()
    }
}

impl std::fmt::Debug for CustomReadObject {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("CustomReadObject")
            .finish_non_exhaustive()
    }
}

impl PartialEq for CustomReadObject {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl Eq for CustomReadObject {}
