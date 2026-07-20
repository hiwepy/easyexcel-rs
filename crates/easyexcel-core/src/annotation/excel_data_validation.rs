//! Mirrors Java `com.alibaba.excel.annotation.write.ExcelDataValidation` (introduced in Phase 1).
//!
//! In Rust, `#[derive(ExcelRow)]` with `#[excel(data_validation(...))]` attribute
//! replaces Java runtime annotation processing. This marker type exists
//! for 1:1 Java file parity.
//!
//! Java original (concept):
//! ```java
//! @Target(ElementType.FIELD)
//! @Retention(RetentionPolicy.RUNTIME)
//! public @interface ExcelDataValidation {
//!     String type() default "list";
//!     String operator() default "between";
//!     String[] formula1() default {};
//!     String[] formula2() default {};
//! }
//! ```

/// Marker type mirroring Java `@ExcelDataValidation`.
///
/// Use in Rust via:
/// ```ignore
/// #[derive(ExcelRow)]
/// struct Demo {
///     #[excel(data_validation(type = "list", formula1 = "A,B,C"))]
///     status: String,
/// }
/// ```
#[allow(dead_code)]
pub struct ExcelDataValidation;