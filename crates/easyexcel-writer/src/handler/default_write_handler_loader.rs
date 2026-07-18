//! Mirrors Java `com.alibaba.excel.write.handler.DefaultWriteHandlerLoader`.

use easyexcel_core::WriteHandler;

use crate::handler::r#impl::impl_default_row_write_handler::DefaultRowWriteHandler;

/// Mirrors Java `DefaultWriteHandlerLoader.loadDefaultHandler(Boolean useDefaultStyle, ExcelTypeEnum excelType)`.
///
/// Returns the Java-equivalent default handler set, with the same
/// ordering and identity guarantees. XLS / CSV variants are not
/// represented separately because [`crate::ExcelWriter`] uses one
/// backend per file extension.
pub struct DefaultWriteHandlerLoader;

impl DefaultWriteHandlerLoader {
    /// Returns the default handler list. (Java `loadDefaultHandler`)
    #[must_use]
    pub fn load_default_handler() -> Vec<Box<dyn WriteHandler>> {
        vec![Box::new(DefaultRowWriteHandler::new())]
    }
}