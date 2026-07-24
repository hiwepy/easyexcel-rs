//! Mirrors Java `com.alibaba.excel.context.xls.*`.

use easyexcel_core::support::ExcelTypeEnum;

use crate::ReadOptions;
use crate::context::read_sheet::ReadSheet;
use crate::holder::xls::xls_read_sheet_holder::XlsReadSheetHolder;
use crate::holder::xls::xls_read_workbook_holder::XlsReadWorkbookHolder;

use super::analysis_context_impl::AnalysisContextImpl;

/// Mirrors Java `XlsReadContext extends AnalysisContext`.
pub trait XlsReadContext {
    /// Returns the shared analysis state.
    fn analysis_context_impl(&self) -> &AnalysisContextImpl;

    /// Returns XLS workbook holder. (Java `xlsReadWorkbookHolder()`)
    fn xls_read_workbook_holder(&self) -> &XlsReadWorkbookHolder;

    /// Returns XLS sheet holder. (Java `xlsReadSheetHolder()`)
    fn xls_read_sheet_holder(&self) -> Option<&XlsReadSheetHolder>;
}

/// Mirrors Java `DefaultXlsReadContext extends AnalysisContextImpl implements XlsReadContext`.
#[derive(Debug, Clone)]
pub struct DefaultXlsReadContext {
    inner: AnalysisContextImpl,
    xls_read_workbook_holder: XlsReadWorkbookHolder,
    xls_read_sheet_holder: Option<XlsReadSheetHolder>,
}

impl DefaultXlsReadContext {
    /// Mirrors Java `DefaultXlsReadContext(ReadWorkbook, ExcelTypeEnum)`.
    #[must_use]
    pub fn new(options: &ReadOptions) -> Self {
        Self {
            inner: AnalysisContextImpl::new(ExcelTypeEnum::Xls, options),
            xls_read_workbook_holder: XlsReadWorkbookHolder::from_options(options),
            xls_read_sheet_holder: None,
        }
    }

    /// Selects the current sheet and materializes the typed XLS holder.
    pub fn current_sheet(&mut self, read_sheet: &ReadSheet) -> easyexcel_core::Result<()> {
        self.inner.current_sheet(read_sheet)?;
        let sheet_no = i32::try_from(read_sheet.sheet_no()).map_err(|_| {
            easyexcel_core::ExcelError::Format("sheet index exceeds i32 range".to_owned())
        })?;
        self.xls_read_sheet_holder =
            Some(XlsReadSheetHolder::new(sheet_no, read_sheet.sheet_name()));
        Ok(())
    }

    /// Returns mutable XLS workbook state for record listeners.
    pub const fn xls_read_workbook_holder_mut(&mut self) -> &mut XlsReadWorkbookHolder {
        &mut self.xls_read_workbook_holder
    }
}

impl XlsReadContext for DefaultXlsReadContext {
    fn analysis_context_impl(&self) -> &AnalysisContextImpl {
        &self.inner
    }

    fn xls_read_workbook_holder(&self) -> &XlsReadWorkbookHolder {
        &self.xls_read_workbook_holder
    }

    fn xls_read_sheet_holder(&self) -> Option<&XlsReadSheetHolder> {
        self.xls_read_sheet_holder.as_ref()
    }
}
