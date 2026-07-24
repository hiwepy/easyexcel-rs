//! Mirrors Java `com.alibaba.excel.context.*`.
//!
//! `default_*_read_context` 文件为 Java 类名 1:1 路径镜像（不删既有实现）。

pub mod analysis_context_impl;
pub mod csv_read_context;
pub mod default_csv_read_context;
pub mod default_xls_read_context;
pub mod default_xlsx_read_context;
pub mod read_sheet;
pub mod xls_read_context;
pub mod xlsx_read_context;

pub use analysis_context_impl::AnalysisContextImpl;
pub use csv_read_context::{CsvReadContext, DefaultCsvReadContext};
pub use read_sheet::ReadSheet;
pub use xls_read_context::{DefaultXlsReadContext, XlsReadContext};
pub use xlsx_read_context::{DefaultXlsxReadContext, XlsxReadContext};

#[cfg(test)]
mod tests {
    use easyexcel_core::support::ExcelTypeEnum;
    use easyexcel_core::{CsvCharset, ExcelError, Result};

    use super::*;
    use crate::ReadOptions;
    use crate::holder::read_workbook_holder::ReadWorkbookHolder;

    fn options() -> ReadOptions {
        ReadOptions {
            charset: CsvCharset::from("GBK"),
            ignore_empty_row: false,
            password: Some("secret".to_owned()),
            ..ReadOptions::default()
        }
    }

    fn assert_workbook_options(holder: &ReadWorkbookHolder) {
        assert_eq!(holder.charset.name(), "GBK");
        assert!(!holder.ignore_empty_row);
        assert_eq!(holder.password.as_deref(), Some("secret"));
        assert!(holder.auto_close_stream);
    }

    #[test]
    fn format_contexts_expose_resolved_workbook_and_current_sheet_holders() -> Result<()> {
        let options = options();
        let sheet = ReadSheet::with_name(2, "Data");

        let mut csv = DefaultCsvReadContext::new(&options);
        assert_eq!(csv.analysis_context_impl().excel_type(), ExcelTypeEnum::Csv);
        assert_workbook_options(csv.csv_read_workbook_holder().inner());
        assert!(csv.csv_read_sheet_holder().is_none());
        csv.current_sheet(&sheet)?;
        let csv_sheet = csv.csv_read_sheet_holder().expect("CSV sheet holder");
        assert_eq!(csv_sheet.inner().sheet_no, 2);
        assert_eq!(csv_sheet.inner().sheet_name, "Data");
        assert_eq!(
            csv.analysis_context_impl().analysis_context().sheet_name(),
            "Data"
        );

        let mut xls = DefaultXlsReadContext::new(&options);
        assert_eq!(xls.analysis_context_impl().excel_type(), ExcelTypeEnum::Xls);
        assert_workbook_options(xls.xls_read_workbook_holder().inner());
        assert!(xls.xls_read_sheet_holder().is_none());
        xls.current_sheet(&sheet)?;
        let xls_sheet = xls.xls_read_sheet_holder().expect("XLS sheet holder");
        assert_eq!(xls_sheet.inner().sheet_no, 2);
        assert_eq!(xls_sheet.inner().sheet_name, "Data");

        let mut xlsx = DefaultXlsxReadContext::new(&options);
        assert_eq!(
            xlsx.analysis_context_impl().excel_type(),
            ExcelTypeEnum::Xlsx
        );
        assert_workbook_options(xlsx.xlsx_read_workbook_holder().inner());
        assert!(xlsx.xlsx_read_sheet_holder().is_none());
        xlsx.current_sheet(&sheet)?;
        let xlsx_sheet = xlsx.xlsx_read_sheet_holder().expect("XLSX sheet holder");
        assert_eq!(xlsx_sheet.inner().sheet_no, 2);
        assert_eq!(xlsx_sheet.inner().sheet_name, "Data");
        Ok(())
    }

    #[test]
    fn typed_current_sheet_preserves_java_duplicate_read_error() {
        let sheet = ReadSheet::with_name(0, "Sheet1");
        let mut csv = DefaultCsvReadContext::new(&ReadOptions::default());
        csv.current_sheet(&sheet).expect("first CSV read");
        let mut xls = DefaultXlsReadContext::new(&ReadOptions::default());
        xls.current_sheet(&sheet).expect("first XLS read");
        let mut xlsx = DefaultXlsxReadContext::new(&ReadOptions::default());
        xlsx.current_sheet(&sheet).expect("first XLSX read");

        for error in [
            csv.current_sheet(&sheet).expect_err("duplicate CSV"),
            xls.current_sheet(&sheet).expect_err("duplicate XLS"),
            xlsx.current_sheet(&sheet).expect_err("duplicate XLSX"),
        ] {
            assert!(matches!(
                error,
                ExcelError::Format(message) if message == "Cannot read sheet repeatedly."
            ));
        }
    }
}
