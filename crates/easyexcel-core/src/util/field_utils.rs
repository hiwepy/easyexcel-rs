#![allow(dead_code)]

//! Mirrors Java com.alibaba.excel.util.FieldUtils.
//!
//! Java uses Spring's `ReflectionUtils` / Apache Commons `FieldUtils` to
//! resolve fields (and to strip CGLIB `$$EnhancerByCGLIB$$` synthetic
//! suffixes). Rust has no runtime reflection, so both helpers are
//! returned as no-op anchors.

/// Mirrors `com.alibaba.excel.util.FieldUtils#resolveCglibFieldName`.
///
/// Java strips the `$$EnhancerByCGLIB$$<hash>` suffix added by the CGLIB
/// proxy. Rust has no equivalent bytecode rewriting, so the input is
/// returned verbatim.
#[must_use]
pub fn resolve_cglib_field_name(name: &str) -> &str {
    name
}

/// Mirrors `com.alibaba.excel.util.FieldUtils#getField`.
///
/// Returns `None` because Rust field access is resolved at compile time
/// via `derive(ExcelRow)` instead of runtime reflection.
#[must_use]
pub fn get_field(_class_name: &str, _field_name: &str) -> Option<()> {
    None
}
