//! Mirrors Java `com.alibaba.excel.write.handler.AbstractWorkbookWriteHandler`.

use easyexcel_core::WriteHandler;

use crate::handler::workbook_write_handler::WorkbookWriteHandler;

/// Mirrors Java `AbstractWorkbookWriteHandler implements WorkbookWriteHandler`.
#[allow(dead_code)]
#[deprecated(note = "Use `easyexcel_core::WriteHandler` directly")]
pub struct AbstractWorkbookWriteHandler;

impl WriteHandler for AbstractWorkbookWriteHandler {
    fn order(&self) -> i32 {
        0
    }
}

impl WorkbookWriteHandler for AbstractWorkbookWriteHandler {}