//! Mirrors Java `com.alibaba.excel.context.csv.*`.

use easyexcel_core::support::ExcelTypeEnum;

use crate::holder::csv::csv_read_sheet_holder::CsvReadSheetHolder;
use crate::holder::csv::csv_read_workbook_holder::CsvReadWorkbookHolder;
use crate::ReadOptions;

use super::analysis_context_impl::AnalysisContextImpl;

/// Mirrors Java `CsvReadContext extends AnalysisContext`.
pub trait CsvReadContext {
    /// Returns the shared analysis state.
    fn analysis_context_impl(&self) -> &AnalysisContextImpl;

    /// Returns CSV workbook holder. (Java `csvReadWorkbookHolder()`)
    fn csv_read_workbook_holder(&self) -> &CsvReadWorkbookHolder;

    /// Returns CSV sheet holder. (Java `csvReadSheetHolder()`)
    fn csv_read_sheet_holder(&self) -> Option<&CsvReadSheetHolder>;
}

/// Mirrors Java `DefaultCsvReadContext extends AnalysisContextImpl implements CsvReadContext`.
#[derive(Debug, Clone)]
pub struct DefaultCsvReadContext {
    inner: AnalysisContextImpl,
    csv_read_workbook_holder: CsvReadWorkbookHolder,
    csv_read_sheet_holder: Option<CsvReadSheetHolder>,
}

impl DefaultCsvReadContext {
    /// Mirrors Java `DefaultCsvReadContext(ReadWorkbook, ExcelTypeEnum)`.
    #[must_use]
    pub fn new(options: &ReadOptions) -> Self {
        Self {
            inner: AnalysisContextImpl::new(ExcelTypeEnum::Csv, options),
            csv_read_workbook_holder: CsvReadWorkbookHolder::new(),
            csv_read_sheet_holder: None,
        }
    }
}

impl CsvReadContext for DefaultCsvReadContext {
    fn analysis_context_impl(&self) -> &AnalysisContextImpl {
        &self.inner
    }

    fn csv_read_workbook_holder(&self) -> &CsvReadWorkbookHolder {
        &self.csv_read_workbook_holder
    }

    fn csv_read_sheet_holder(&self) -> Option<&CsvReadSheetHolder> {
        self.csv_read_sheet_holder.as_ref()
    }
}
