//! Mirrors Java `com.alibaba.excel.metadata.csv.CsvWorkbook`.

use crate::CsvCharset;
use crate::excel_error::ExcelError;
use crate::util::work_book_util::SheetCreator;

use super::csv_cell_style::CsvCellStyle;
use super::csv_data_format::CsvDataFormat;
use super::csv_sheet::CsvSheet;

/// Logical workbook used by the streaming CSV backend.
#[derive(Debug, Clone, PartialEq)]
pub struct CsvWorkbook {
    locale: String,
    use_1904_windowing: bool,
    use_scientific_format: bool,
    charset: CsvCharset,
    with_bom: bool,
    sheet: Option<CsvSheet>,
    data_format: CsvDataFormat,
    cell_styles: Vec<CsvCellStyle>,
}

impl CsvWorkbook {
    /// Creates a CSV workbook with its global rendering options.
    #[must_use]
    pub fn new(
        locale: impl Into<String>,
        use_1904_windowing: bool,
        use_scientific_format: bool,
        charset: CsvCharset,
        with_bom: bool,
    ) -> Self {
        Self {
            locale: locale.into(),
            use_1904_windowing,
            use_scientific_format,
            charset,
            with_bom,
            sheet: None,
            data_format: CsvDataFormat::new(),
            cell_styles: Vec::new(),
        }
    }

    /// Returns the configured locale tag.
    #[must_use]
    pub fn locale(&self) -> &str {
        &self.locale
    }

    /// Returns the configured charset.
    #[must_use]
    pub const fn charset(&self) -> &CsvCharset {
        &self.charset
    }

    /// Returns whether output starts with a charset BOM.
    #[must_use]
    pub const fn with_bom(&self) -> bool {
        self.with_bom
    }

    /// Returns whether the 1904 date system is enabled.
    #[must_use]
    pub const fn use_1904_windowing(&self) -> bool {
        self.use_1904_windowing
    }

    /// Returns whether large/small numbers use scientific notation.
    #[must_use]
    pub const fn use_scientific_format(&self) -> bool {
        self.use_scientific_format
    }

    /// Returns the only CSV sheet, when it has been created.
    #[must_use]
    pub const fn sheet(&self) -> Option<&CsvSheet> {
        self.sheet.as_ref()
    }

    /// Returns the workbook-local data-format registry.
    pub const fn data_format_mut(&mut self) -> &mut CsvDataFormat {
        &mut self.data_format
    }

    /// Creates and registers a cell style.
    pub fn create_cell_style(&mut self) -> &mut CsvCellStyle {
        let index = i16::try_from(self.cell_styles.len()).unwrap_or(i16::MAX);
        self.cell_styles.push(CsvCellStyle::new(index));
        self.cell_styles.last_mut().expect("just pushed")
    }

    /// Returns a registered cell style.
    #[must_use]
    pub fn cell_style(&self, index: usize) -> Option<&CsvCellStyle> {
        self.cell_styles.get(index)
    }
}

impl SheetCreator for CsvWorkbook {
    type Sheet<'a>
        = &'a mut CsvSheet
    where
        Self: 'a;

    fn create_sheet(&mut self, sheet_name: &str) -> Result<Self::Sheet<'_>, ExcelError> {
        if self.sheet.is_some() {
            return Err(ExcelError::Unsupported(
                "CSV repeat sheet creation is not allowed".to_owned(),
            ));
        }
        self.sheet = Some(CsvSheet::new(sheet_name));
        Ok(self.sheet.as_mut().expect("just assigned"))
    }
}
