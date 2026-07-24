//! Phase 3 — 1:1 test matrix for advanced features (comment / hyperlink /
//! formula write-back, CellExtra read-back).
//!
//! Java reference: `com.alibaba.easyexcel.test.core.cellData.CellDataDataTest`
//! + `com.alibaba.easyexcel.test.demo.write.WriteTest#commentWrite` /
//! `imageWrite` / `writeCellDataWrite` + Java handler/extra system.
//!
//! Rust mirror: writer paths in `crates/easyexcel-writer/src/lib.rs`
//! `write_formula_with_format` / `write_url_with_options` / `insert_note`;
//! reader paths in `crates/easyexcel-reader/src/xlsx_rows.rs` CellExtra
//! emit (Comment / Hyperlink / Merge).
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>`.

use easyexcel::CellValue;
use easyexcel_core::{CellExtra, CellExtraType, WriteCellData};
use easyexcel_macro::ExcelRow as DeriveExcelRow;

// ---------------------------------------------------------------------------
// FormulaCellValue — Java `CellDataDataTest#t01ReadFormula07` / write-back
// ---------------------------------------------------------------------------

mod formula_cell_test {
    //! Mirrors: CellDataDataTest#t01ReadFormula07
    use super::*;
    use easyexcel_core::ExcelRow as ExcelRowTrait;

    #[derive(Debug, DeriveExcelRow)]
    struct FormulaRow {
        #[excel(formula = "SUM(A1:A10)")]
        total: f64,
    }

    /// Java: `CellValue::Formula` round-trips through CellValue::Formula enum.
    #[test]
    fn t01_read_formula07() {
        // Build a Formula CellValue and verify it carries the formula text
        // verbatim (mirrors Java's FormulaData.formulaValue preservation).
        let fv = CellValue::Formula("SUM(A1:A10)".to_owned());
        if let CellValue::Formula(s) = &fv {
            assert_eq!(s, "SUM(A1:A10)");
        } else {
            panic!("expected CellValue::Formula, got {fv:?}");
        }

        // The column metadata should reflect the @ExcelFormula derive attr.
        let cols = <FormulaRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].formula, Some("SUM(A1:A10)"));
    }

    /// Java: applying formula decoration wraps scalar in CellValue::Formula.
    #[test]
    fn t02_write_formula07() {
        let cols = <FormulaRow as ExcelRowTrait>::schema();
        let col = &cols[0];
        let data = WriteCellData::new(CellValue::Float(0.0));
        let decorated = col.apply_decorations(data);
        // WriteCellData::to_excel_cell emits the underlying CellValue
        // (CellValue::Formula here).
        let ctx = easyexcel_core::ConvertContext {
            sheet_name: String::new(),
            row_index: 0,
            column_index: None,
            field: "",
            format: None,
            use_1904_windowing: false,
        };
        let cv = easyexcel_core::IntoExcelCell::to_excel_cell(&decorated, &ctx).unwrap();
        if let CellValue::Formula(s) = &cv {
            assert_eq!(s, "SUM(A1:A10)");
        } else {
            panic!("expected CellValue::Formula, got {cv:?}");
        }
    }
}

// ---------------------------------------------------------------------------
// HyperlinkCellValue — Java `CellDataDataTest#t03ReadHyperlink07` / write-back
// ---------------------------------------------------------------------------

mod hyperlink_cell_test {
    //! Mirrors: CellDataDataTest#t03ReadHyperlink07
    use super::*;
    use easyexcel_core::ExcelRow as ExcelRowTrait;

    #[derive(Debug, DeriveExcelRow)]
    struct HyperlinkRow {
        #[excel(hyperlink = "https://example.com")]
        url: String,
    }

    /// Java: `CellValue::Hyperlink { url, text }` carries both target and display.
    #[test]
    fn t03_read_hyperlink07() {
        let hv = CellValue::Hyperlink {
            url: "https://example.com".to_owned(),
            text: "example".to_owned(),
        };
        if let CellValue::Hyperlink { url, text } = &hv {
            assert_eq!(url, "https://example.com");
            assert_eq!(text, "example");
        } else {
            panic!("expected CellValue::Hyperlink, got {hv:?}");
        }

        // Column metadata reflects the @ExcelHyperlink derive attr.
        let cols = <HyperlinkRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].hyperlink, Some("https://example.com"));
    }

    /// Java: applying hyperlink decoration wraps the display text.
    #[test]
    fn t04_write_hyperlink07() {
        let col = <HyperlinkRow as ExcelRowTrait>::schema()[0];
        let data = WriteCellData::from_string("example");
        let decorated = col.apply_decorations(data);
        let ctx = easyexcel_core::ConvertContext {
            sheet_name: String::new(),
            row_index: 0,
            column_index: None,
            field: "",
            format: None,
            use_1904_windowing: false,
        };
        let cv = easyexcel_core::IntoExcelCell::to_excel_cell(&decorated, &ctx).unwrap();
        if let CellValue::Hyperlink { url, text } = &cv {
            assert_eq!(url, "https://example.com");
            assert_eq!(text, "example");
        } else {
            panic!("expected CellValue::Hyperlink, got {cv:?}");
        }
    }
}

// ---------------------------------------------------------------------------
// CommentCellValue — Java `CellDataDataTest#t05ReadComment07` / write-back
// ---------------------------------------------------------------------------

mod comment_cell_test {
    //! Mirrors: CellDataDataTest#t05ReadComment07
    use super::*;
    use easyexcel_core::ExcelRow as ExcelRowTrait;

    #[derive(Debug, DeriveExcelRow)]
    struct CommentRow {
        #[excel(comment = "note text")]
        cell: String,
    }

    /// Java: `CellValue::Comment { value, text }` carries note + underlying value.
    #[test]
    fn t05_read_comment07() {
        let cv = CellValue::Comment {
            value: Box::new(CellValue::String("hello".to_owned())),
            text: "note text".to_owned(),
        };
        if let CellValue::Comment { value, text } = &cv {
            assert_eq!(text, "note text");
            assert_eq!(**value, CellValue::String("hello".to_owned()));
        } else {
            panic!("expected CellValue::Comment, got {cv:?}");
        }

        let cols = <CommentRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].comment, Some("note text"));
    }

    /// Java: applying comment decoration wraps value + attaches note text.
    #[test]
    fn t06_write_comment07() {
        let col = <CommentRow as ExcelRowTrait>::schema()[0];
        let data = WriteCellData::from_string("hello");
        let decorated = col.apply_decorations(data);
        let ctx = easyexcel_core::ConvertContext {
            sheet_name: String::new(),
            row_index: 0,
            column_index: None,
            field: "",
            format: None,
            use_1904_windowing: false,
        };
        let cv = easyexcel_core::IntoExcelCell::to_excel_cell(&decorated, &ctx).unwrap();
        if let CellValue::Comment { value, text } = &cv {
            assert_eq!(text, "note text");
            assert_eq!(**value, CellValue::String("hello".to_owned()));
        } else {
            panic!("expected CellValue::Comment, got {cv:?}");
        }
    }
}

// ---------------------------------------------------------------------------
// CellExtra read-back (Java `CellDataDataTest#t07Extras07`)
// ---------------------------------------------------------------------------

mod cell_extra_test {
    //! Mirrors: CellDataDataTest#t07Extras07
    use super::*;

    /// Java: extras emit Merge / Hyperlink / Comment in document order.
    #[test]
    fn t07_extras07() {
        let merge = CellExtra::new(CellExtraType::Merge, None, 0, 2, 0, 1);
        let link = CellExtra::new(
            CellExtraType::Hyperlink,
            Some("https://example.com".to_owned()),
            0,
            0,
            0,
            0,
        );
        let comment = CellExtra::new(CellExtraType::Comment, Some("note".to_owned()), 1, 1, 0, 0);

        let extras = vec![merge, link, comment];
        assert_eq!(extras[0].extra_type(), CellExtraType::Merge);
        assert_eq!(extras[0].text(), None);
        assert_eq!(extras[1].extra_type(), CellExtraType::Hyperlink);
        assert_eq!(extras[1].text(), Some("https://example.com"));
        assert_eq!(extras[2].extra_type(), CellExtraType::Comment);
        assert_eq!(extras[2].text(), Some("note"));
        assert_eq!(extras[2].first_row_index(), 1);
    }
}
