//! Mirrors Java `com.alibaba.excel.write.handler.AbstractRowWriteHandler`.

use easyexcel_core::WriteHandler;

use crate::handler::row_write_handler::RowWriteHandler;

/// Mirrors Java `AbstractRowWriteHandler implements RowWriteHandler`.
#[allow(dead_code)]
#[deprecated(note = "Use `easyexcel_core::WriteHandler` directly")]
pub struct AbstractRowWriteHandler;

impl WriteHandler for AbstractRowWriteHandler {
    fn order(&self) -> i32 {
        0
    }
}

impl RowWriteHandler for AbstractRowWriteHandler {}