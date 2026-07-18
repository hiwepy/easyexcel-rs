//! Mirrors Java `com.alibaba.excel.write.handler.AbstractCellWriteHandler`.

use easyexcel_core::WriteCellContext;
use easyexcel_core::WriteHandler;

use crate::handler::cell_write_handler::CellWriteHandler;

/// Mirrors Java `AbstractCellWriteHandler implements CellWriteHandler`.
///
/// Java declares the type as `@Deprecated`; Rust keeps the same
/// name and delegates the three callbacks to default no-ops so older
/// user code that imports it still compiles.
#[allow(dead_code)]
#[deprecated(note = "Use `easyexcel_core::WriteHandler` directly")]
pub struct AbstractCellWriteHandler;

impl WriteHandler for AbstractCellWriteHandler {
    fn order(&self) -> i32 {
        0
    }
}

impl CellWriteHandler for AbstractCellWriteHandler {
    // All three callbacks remain no-ops — the trait provides sensible
    // defaults; we just need a concrete type for the deprecated shim.
}