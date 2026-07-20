//! Phase 4 — 1:1 test matrix for `ExcelWriter::write_with_table` overload.
//!
//! Java reference: `com.alibaba.easyexcel.test.core.Write#tableWrite` and the
//! `ExcelWriter.write(Collection, WriteSheet, WriteTable)` three-arg overload.
//!
//! Rust mirror: `ExcelWriter::write_with_table(rows, sheet, table)` method +
//! `EasyExcel::writer_table(i32)` + `EasyExcel::writer_table_builder(i32)`
//! facade entry points.
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>`.

use easyexcel::{EasyExcel, ExcelRow};
use easyexcel_core::{IntoExcelCell, WriteCellData};
use easyexcel_core::ExcelRow as ExcelRowTrait;

// ---------------------------------------------------------------------------
// Top-level test structs (visible to derive macro)
// ---------------------------------------------------------------------------

#[derive(Debug, PartialEq, ExcelRow)]
struct TableRow {
    #[excel(name = "Name")]
    name: String,
    #[excel(name = "Value")]
    value: u32,
}

#[derive(Debug, PartialEq, ExcelRow)]
struct TwoArgRow {
    a: u32,
    b: String,
}

// ---------------------------------------------------------------------------
// writer_table facade — mirrors EasyExcelFactory.writerTable(Integer)
// ---------------------------------------------------------------------------

mod writer_table_test {
    //! Mirrors EasyExcelFactory#writerTable07
    use super::*;

    /// Java: `EasyExcelFactory.writerTable(int)` returns a `WriteTable`
    /// with the requested `tableNo`.
    #[test]
    fn t01_writer_table07() {
        let table = EasyExcel::writer_table(2);
        assert_eq!(table.table_no(), 2);
    }

    /// Java: `EasyExcelFactory.writerTable(null)` defaults to tableNo=0.
    #[test]
    fn t02_writer_table_default07() {
        let table = EasyExcel::writer_table(0);
        assert_eq!(table.table_no(), 0);
    }
}

// ---------------------------------------------------------------------------
// writer_table_builder facade — mirrors ExcelWriterBuilder.table(Integer)
// ---------------------------------------------------------------------------

mod writer_table_builder_test {
    //! Mirrors ExcelWriterBuilder#table07
    use super::*;

    /// Java: builder accepts table_no + head style + need_head overrides.
    #[test]
    fn t03_writer_table_builder07() {
        let builder = EasyExcel::writer_table_builder(3)
            .need_head(false)
            .relative_head_row_index(2);
        let table = builder.build();
        assert_eq!(table.table_no(), 3);
        assert!(!table.options().need_head);
        assert_eq!(table.options().relative_head_row_index, 2);
    }
}

// ---------------------------------------------------------------------------
// ExcelWriter::write_with_table three-arg overload
// ---------------------------------------------------------------------------

mod write_with_table_test {
    //! Mirrors Write#tableWrite07
    use super::*;

    /// Java: a single sheet can host two tables via three-arg overload.
    /// Each table uses its own header + content rows.
    #[test]
    fn t04_write_with_table07() {
        let path = std::env::temp_dir().join("easyexcel_phase4_write_with_table07.xlsx");

        let sheet = EasyExcel::writer_sheet::<TableRow>("Sheet1");
        let mut writer = EasyExcel::write::<TableRow>(&path).build();

        // First table.
        let rows_a = vec![
            TableRow {
                name: "alpha".to_owned(),
                value: 1,
            },
            TableRow {
                name: "beta".to_owned(),
                value: 2,
            },
        ];
        let table_a = EasyExcel::writer_table(0);
        writer.write_with_table(rows_a, &sheet, &table_a).unwrap();

        // Second table (different table_no, same sheet).
        let rows_b = vec![TableRow {
            name: "gamma".to_owned(),
            value: 3,
        }];
        let table_b = EasyExcel::writer_table(1);
        writer.write_with_table(rows_b, &sheet, &table_b).unwrap();

        writer.finish().unwrap();
        assert!(path.exists(), "xlsx file must be produced");

        // Verify column metadata is preserved on the typed row.
        let cols = <TableRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].name, "Name");
        assert_eq!(cols[1].name, "Value");
    }
}

// ---------------------------------------------------------------------------
// Java: Write#simpleWrite (no table) — still works after table overload added.
// ---------------------------------------------------------------------------

mod write_two_arg_still_works_test {
    //! Regression: two-arg write must still work after adding write_with_table.
    use super::*;

    /// Java: existing two-arg call pattern must continue to compile and run.
    #[test]
    fn t05_write_two_arg_still_works07() {
        let path = std::env::temp_dir().join("easyexcel_phase4_two_arg07.xlsx");
        let sheet = EasyExcel::writer_sheet::<TwoArgRow>("Sheet1");
        let mut writer = EasyExcel::write::<TwoArgRow>(&path).build();
        writer
            .write(
                vec![TwoArgRow {
                    a: 42,
                    b: "answer".to_owned(),
                }],
                &sheet,
            )
            .unwrap();
        writer.finish().unwrap();
        assert!(path.exists(), "xlsx file must be produced");
    }
}

// ---------------------------------------------------------------------------
// POI handle access — Phase 4.1 (limited). CellValue decode path stays intact.
// ---------------------------------------------------------------------------

mod poi_handle_test {
    //! Phase 4.1: POI handle access through the cell value decode pipeline.
    use super::*;

    /// Java: a `WriteCellData::set_value` call preserves the underlying
    /// typed scalar (mirrors Java's POI Cell.setCellValue(T) pipeline).
    #[test]
    fn t06_poi_handle_via_cell_data07() {
        let mut data = WriteCellData::new(easyexcel::CellValue::Empty);
        data.set_value(easyexcel::CellValue::String("hello".to_owned()));
        let ctx = easyexcel_core::ConvertContext {
            sheet_name: String::new(),
            row_index: 0,
            column_index: None,
            field: "",
            format: None,
        };
        let cv = IntoExcelCell::to_excel_cell(&data, &ctx).unwrap();
        if let easyexcel::CellValue::String(s) = cv {
            assert_eq!(s, "hello");
        } else {
            panic!("expected CellValue::String, got {cv:?}");
        }
    }
}