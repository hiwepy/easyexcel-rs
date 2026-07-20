//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelFormula` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(formula = "...")]` attribute
//! replaces Java runtime annotation processing. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelFormula {
//!     String value();
//! }
//! ```

/// Marker type mirroring Java `@ExcelFormula`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(formula = "SUM(A1:A10)")]
///     total: f64,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelFormula;