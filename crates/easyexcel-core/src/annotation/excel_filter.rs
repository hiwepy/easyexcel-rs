//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelFilter` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(filter)]` attribute
//! enables auto-filtering on the column. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelFilter {}
//! ```

/// Marker type mirroring Java `@ExcelFilter`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(filter)]
///     name: String,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelFilter;
