//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelComment` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(comment = "...")]` attribute
//! replaces Java runtime annotation processing. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelComment {
//!     String value();
//!     String author() default "";
//! }
//! ```

/// Marker type mirroring Java `@ExcelComment`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(comment = "TODO: validate this row")]
///     note: String,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelComment;