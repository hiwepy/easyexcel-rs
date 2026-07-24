//! Mirrors Java `com.alibaba.excel.analysis.ExcelReadExecutor` (interface).

use std::path::PathBuf;

use easyexcel_core::support::ExcelTypeEnum;
use easyexcel_core::{ExcelRow, ReadListener, Result};

use crate::ReadOptions;
use crate::analysis::csv::csv_excel_read_executor::CsvExcelReadExecutor;
use crate::analysis::v03::xls_sax_analyser::XlsSaxAnalyser;
use crate::analysis::v07::xlsx_sax_analyser::XlsxSaxAnalyser;
use crate::context::ReadSheet;

/// Mirrors Java `ExcelReadExecutor`.
///
/// Java declares `sheetList()` and `execute()`. Rust's `read_xlsx` /
/// `read_xls` / `read_csv` functions cover the same contract.
pub trait ExcelReadExecutor {
    /// Returns discovered worksheets. (Java `sheetList()`)
    fn sheet_list(&self) -> &[ReadSheet];

    /// Executes the read with Rust's typed listener and current options.
    ///
    /// Java retrieves erased listeners and sheet parameters from
    /// `ReadWorkbook`; Rust makes those dependencies explicit.
    fn execute<T, L>(&mut self, options: &ReadOptions, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>;
}

/// Runtime executor selected by `ExcelAnalyserImpl.choiceExcelExecutor`.
///
/// Java stores one `ExcelReadExecutor` interface object. Rust uses an enum so
/// callers can inspect the same concrete XLSX/XLS/CSV executor that owns sheet
/// discovery while typed listener execution remains statically dispatched.
pub enum ExcelReadExecutorKind {
    /// OOXML SAX executor.
    Xlsx(XlsxSaxAnalyser),
    /// BIFF event executor.
    Xls(XlsSaxAnalyser),
    /// CSV record executor.
    Csv(CsvExcelReadExecutor),
}

impl ExcelReadExecutorKind {
    /// Constructs the concrete executor selected from the resolved workbook type.
    pub fn new(
        excel_type: ExcelTypeEnum,
        path: impl Into<PathBuf>,
        options: ReadOptions,
    ) -> Result<Self> {
        let path = path.into();
        match excel_type {
            ExcelTypeEnum::Xlsx => XlsxSaxAnalyser::from_path(path, options).map(Self::Xlsx),
            ExcelTypeEnum::Xls => XlsSaxAnalyser::from_path(path, options).map(Self::Xls),
            ExcelTypeEnum::Csv => Ok(Self::Csv(CsvExcelReadExecutor::from_path(path))),
        }
    }

    /// Executes through the selected real parser with the current analyser options.
    pub fn execute_with_listener<T, L>(
        &mut self,
        options: &ReadOptions,
        listener: &mut L,
    ) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        ExcelReadExecutor::execute::<T, L>(self, options, listener)
    }

    /// Returns the concrete executor variant's resolved workbook type.
    #[must_use]
    pub const fn excel_type(&self) -> ExcelTypeEnum {
        match self {
            Self::Xlsx(_) => ExcelTypeEnum::Xlsx,
            Self::Xls(_) => ExcelTypeEnum::Xls,
            Self::Csv(_) => ExcelTypeEnum::Csv,
        }
    }

    /// Returns the selected executor's discovered worksheet list.
    #[must_use]
    pub fn sheet_list(&self) -> &[ReadSheet] {
        ExcelReadExecutor::sheet_list(self)
    }
}

impl ExcelReadExecutor for ExcelReadExecutorKind {
    fn sheet_list(&self) -> &[ReadSheet] {
        match self {
            Self::Xlsx(executor) => executor.sheet_list(),
            Self::Xls(executor) => executor.sheet_list(),
            Self::Csv(executor) => executor.sheet_list(),
        }
    }

    fn execute<T, L>(&mut self, options: &ReadOptions, listener: &mut L) -> Result<()>
    where
        T: ExcelRow,
        L: ReadListener<T>,
    {
        match self {
            Self::Xlsx(executor) => executor.execute::<T, L>(options, listener),
            Self::Xls(executor) => executor.execute::<T, L>(options, listener),
            Self::Csv(executor) => executor.execute::<T, L>(options, listener),
        }
    }
}
