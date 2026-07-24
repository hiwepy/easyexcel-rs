//! Mirrors Java `com.alibaba.excel.context.xlsx.*`.

use easyexcel_core::support::ExcelTypeEnum;

use crate::ReadOptions;
use crate::context::read_sheet::ReadSheet;
use crate::holder::xlsx::xlsx_read_sheet_holder::XlsxReadSheetHolder;
use crate::holder::xlsx::xlsx_read_workbook_holder::XlsxReadWorkbookHolder;

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
            xlsx_read_workbook_holder: XlsxReadWorkbookHolder::from_options(options),
            xlsx_read_sheet_holder: None,
        }
    }

    /// Selects the current sheet and materializes the typed XLSX holder.
    pub fn current_sheet(&mut self, read_sheet: &ReadSheet) -> easyexcel_core::Result<()> {
        self.inner.current_sheet(read_sheet)?;
        let sheet_no = i32::try_from(read_sheet.sheet_no()).map_err(|_| {
            easyexcel_core::ExcelError::Format("sheet index exceeds i32 range".to_owned())
        })?;
        self.xlsx_read_sheet_holder =
            Some(XlsxReadSheetHolder::new(sheet_no, read_sheet.sheet_name()));
        Ok(())
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
