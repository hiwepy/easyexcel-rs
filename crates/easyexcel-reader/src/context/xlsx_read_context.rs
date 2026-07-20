//! Mirrors Java `com.alibaba.excel.context.xlsx.*`.

use easyexcel_core::support::ExcelTypeEnum;

use crate::holder::xlsx::xlsx_read_sheet_holder::XlsxReadSheetHolder;
use crate::holder::xlsx::xlsx_read_workbook_holder::XlsxReadWorkbookHolder;
use crate::ReadOptions;

use super::analysis_context_impl::AnalysisContextImpl;

/// Mirrors Java `XlsxReadContext extends AnalysisContext`.
pub trait XlsxReadContext {
    /// Returns the shared analysis state. (Java `AnalysisContext` methods)
    fn analysis_context_impl(&self) -> &AnalysisContextImpl;

    /// Returns XLSX workbook holder. (Java `xlsxReadWorkbookHolder()`)
    fn xlsx_read_workbook_holder(&self) -> &XlsxReadWorkbookHolder;

    /// Returns XLSX sheet holder. (Java `xlsxReadSheetHolder()`)
    fn xlsx_read_sheet_holder(&self) -> Option<&XlsxReadSheetHolder>;
}

/// Mirrors Java `DefaultXlsxReadContext extends AnalysisContextImpl implements XlsxReadContext`.
#[derive(Debug, Clone)]
pub struct DefaultXlsxReadContext {
    /// Shared analysis state.
    inner: AnalysisContextImpl,
    /// XLSX workbook holder.
    xlsx_read_workbook_holder: XlsxReadWorkbookHolder,
    /// Active XLSX sheet holder.
    xlsx_read_sheet_holder: Option<XlsxReadSheetHolder>,
}

impl DefaultXlsxReadContext {
    /// Mirrors Java `DefaultXlsxReadContext(ReadWorkbook, ExcelTypeEnum)`.
    #[must_use]
    pub fn new(options: &ReadOptions) -> Self {
        Self {
            inner: AnalysisContextImpl::new(ExcelTypeEnum::Xlsx, options),
            xlsx_read_workbook_holder: XlsxReadWorkbookHolder::new(),
            xlsx_read_sheet_holder: None,
        }
    }
}

impl XlsxReadContext for DefaultXlsxReadContext {
    fn analysis_context_impl(&self) -> &AnalysisContextImpl {
        &self.inner
    }

    fn xlsx_read_workbook_holder(&self) -> &XlsxReadWorkbookHolder {
        &self.xlsx_read_workbook_holder
    }

    fn xlsx_read_sheet_holder(&self) -> Option<&XlsxReadSheetHolder> {
        self.xlsx_read_sheet_holder.as_ref()
    }
}
