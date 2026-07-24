//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelConditional` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(conditional(...))]` attribute
//! replaces Java runtime annotation processing. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelConditional {
//!     String condition();        // e.g. "greaterThan(100)"
//!     String fontColor() default "";
//!     String backgroundColor() default "";
//! }
//! ```

/// Marker type mirroring Java `@ExcelConditional`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(conditional(condition = "greaterThan(100)", font_color = "red"))]
///     value: f64,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelConditional;
