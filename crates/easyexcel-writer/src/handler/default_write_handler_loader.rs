//! Mirrors Java `com.alibaba.excel.write.handler.DefaultWriteHandlerLoader`.

use easyexcel_core::WriteHandler;

use crate::handler::r#impl::impl_default_row_write_handler::DefaultRowWriteHandler;
use crate::handler::r#impl::impl_default_sheet_write_handler::DefaultWriteSheetHandler;
use crate::handler::r#impl::impl_default_workbook_write_handler::DefaultWriteWorkbookHandler;

/// Mirrors Java `DefaultWriteHandlerLoader.loadDefaultHandler(Boolean useDefaultStyle, ExcelTypeEnum excelType)`.
///
/// Returns the Java-equivalent default handler set, with the same
/// ordering and identity guarantees. XLS / CSV variants are not
/// represented separately because [`crate::ExcelWriter`] uses one
/// backend per file extension.
pub struct DefaultWriteHandlerLoader;

impl DefaultWriteHandlerLoader {
    /// Returns the default handler list. (Java `loadDefaultHandler`)
    ///
    /// Ordering matches Java `ExcelBuilderImpl.initHandlerChain`:
    /// 1. `DefaultWriteWorkbookHandler` (workbook-level tracking)
    /// 2. `DefaultRowWriteHandler` (freeze-head / row metadata)
    /// 3. `DefaultWriteSheetHandler` (sheet initialization marker)
    #[must_use]
    pub fn load_default_handler() -> Vec<Box<dyn WriteHandler>> {
        vec![
            Box::new(DefaultWriteWorkbookHandler::new()),
            Box::new(DefaultRowWriteHandler::new()),
            Box::new(DefaultWriteSheetHandler::new()),
        ]
    }
}