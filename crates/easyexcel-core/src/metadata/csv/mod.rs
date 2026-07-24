//! Mirrors Java `com.alibaba.excel.metadata.csv.*`.

pub mod csv_cell;
pub mod csv_cell_style;
pub mod csv_data_format;
pub mod csv_rich_text_string;
pub mod csv_row;
pub mod csv_sheet;
pub mod csv_workbook;

pub use csv_cell::CsvCell;
pub use csv_cell_style::CsvCellStyle;
pub use csv_data_format::CsvDataFormat;
pub use csv_rich_text_string::CsvRichTextString;
pub use csv_row::CsvRow;
pub use csv_sheet::CsvSheet;
pub use csv_workbook::CsvWorkbook;

#[cfg(test)]
mod tests {
    use crate::util::work_book_util::{create_cell, create_row, create_sheet};
    use crate::{CellValue, CsvCharset, ExcelError, NumericCellType};

    use super::*;

    #[test]
    fn workbook_sheet_row_cell_chain_preserves_typed_values() {
        let mut workbook = CsvWorkbook::new("zh-CN", false, true, CsvCharset::utf8(), true);
        let sheet = create_sheet(&mut workbook, "用户").expect("sheet");
        let row = create_row(sheet, 0).expect("row");
        create_cell(row, 2)
            .expect("cell")
            .set_value(CellValue::Int(42));
        create_cell(row, 0)
            .expect("cell")
            .set_value(CellValue::String("张三".to_owned()));
        assert_eq!(
            row.cell(2).and_then(CsvCell::numeric_cell_type),
            Some(NumericCellType::Number)
        );
        let record = sheet.take_last_row().expect("row").into_record(3);
        assert_eq!(record, ["张三", "", "42"]);
    }

    #[test]
    fn csv_enforces_single_sheet_ordered_rows_and_unique_cells() {
        let mut workbook = CsvWorkbook::new("und", false, false, CsvCharset::utf8(), false);
        let sheet = create_sheet(&mut workbook, "Sheet1").expect("sheet");
        let row = create_row(sheet, 0).expect("row");
        create_cell(row, 0).expect("cell");
        assert!(matches!(
            create_cell(row, 0),
            Err(ExcelError::Format(message)) if message.contains("already exists")
        ));
        assert!(matches!(
            create_row(sheet, 2),
            Err(ExcelError::Format(message)) if message.contains("expected 1")
        ));
        assert!(matches!(
            create_sheet(&mut workbook, "Sheet2"),
            Err(ExcelError::Unsupported(message)) if message.contains("repeat")
        ));
    }
}
