//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelHyperlink` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(hyperlink = "...")]` attribute
//! replaces Java runtime annotation processing. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelHyperlink {
//!     String value();          // URL or internal location
//!     String display() default "";
//! }
//! ```

/// Marker type mirroring Java `@ExcelHyperlink`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(hyperlink = "https://example.com")]
///     url: String,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelHyperlink;
