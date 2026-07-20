//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelImage` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(image = "...")]` attribute
//! replaces Java runtime annotation processing. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelImage {
//!     String image() default "";
//! }
//! ```

/// Marker type mirroring Java `@ExcelImage`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(image = "tests/fixtures/converter/img.jpg")]
///     logo: Vec<u8>,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelImage;