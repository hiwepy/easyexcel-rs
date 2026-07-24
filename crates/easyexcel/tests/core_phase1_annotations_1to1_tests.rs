//! Phase 1 — 1:1 test matrix for new annotation markers.
//!
//! Java reference: `com.alibaba.easyexcel.test.core.annotation.AnnotationData`
//! extension methods t07..t14 for the seven new annotation markers:
//! `Excel{Image,Comment,Hyperlink,Formula,DataValidation,Conditional,Filter}`.
//!
//! Rust mirror: `#[derive(ExcelRow)]` + `#[excel(image/comment/hyperlink/
//! formula/data_validation/conditional/filter)]` derive attributes that
//! populate `ExcelColumn` and `ExcelWriteMetadata`.
//!
//! These tests verify the metadata is correctly attached to the column
//! at compile-time and runtime. Phase 3 will exercise the actual write
//! path (image insertion, comment XML emit, hyperlink + relationship, etc.).
//!
//! Naming follows: `mod <java_class_snake>` + `fn <java_method_snake>`.

use easyexcel::{CellValue, WriteCellData};
// ExcelRow trait is aliased as ExcelRowTrait to avoid name clash with the
// `easyexcel_derive::ExcelRow` derive macro.
use easyexcel_core::ExcelRow as ExcelRowTrait;

// ---------------------------------------------------------------------------
// AnnotationData (mirrors com.alibaba.easyexcel.test.core.annotation.AnnotationData)
// ---------------------------------------------------------------------------

mod annotation_phase1_image_test {
    //! Mirrors: AnnotationDataTest#t07ExcelImage07
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, PartialEq, ExcelRow)]
    struct ImageData {
        #[excel(image = "tests/fixtures/converter/img.jpg")]
        logo: Vec<u8>,
        name: String,
    }

    /// Java: `@ExcelImage` attribute should be present in the column metadata.
    #[test]
    fn t07_excel_image07() {
        let cols = <ImageData as ExcelRowTrait>::schema();
        assert_eq!(cols.len(), 2);
        assert_eq!(cols[0].field, "logo");
        assert_eq!(cols[0].image_path, Some("tests/fixtures/converter/img.jpg"));
        assert!(cols[1].image_path.is_none());
    }
}

mod annotation_phase1_comment_test {
    //! Mirrors: AnnotationDataTest#t08ExcelComment07
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, PartialEq, ExcelRow)]
    struct CommentRow {
        #[excel(comment = "TODO: validate")]
        note: String,
        count: u32,
    }

    /// Java: `@ExcelComment` attribute should populate the column comment field.
    #[test]
    fn t08_excel_comment07() {
        let cols = <CommentRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].comment, Some("TODO: validate"));
        assert!(cols[1].comment.is_none());
    }
}

mod annotation_phase1_hyperlink_test {
    //! Mirrors: AnnotationDataTest#t09ExcelHyperlink07
    use easyexcel::{CellValue, WriteCellData};
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct HyperlinkRow {
        #[excel(hyperlink = "https://example.com")]
        url: String,
        label: String,
    }

    /// Java: `@ExcelHyperlink` attribute should populate the column hyperlink field.
    #[test]
    fn t09_excel_hyperlink07() {
        let cols = <HyperlinkRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].hyperlink, Some("https://example.com"));
        assert!(cols[1].hyperlink.is_none());
    }

    /// Java: applying decoration should produce a Hyperlink CellValue.
    #[test]
    fn t09_apply_decoration07() {
        let col = <HyperlinkRow as ExcelRowTrait>::schema()[0];
        let data = WriteCellData::from_string("Click here");
        let decorated = col.apply_decorations(data);
        match decorated.value() {
            CellValue::Hyperlink { url, text } => {
                assert_eq!(url, "https://example.com");
                assert_eq!(text, "Click here");
            }
            other => panic!("expected CellValue::Hyperlink, got {other:?}"),
        }
    }
}

mod annotation_phase1_formula_test {
    //! Mirrors: AnnotationDataTest#t10ExcelFormula07
    use easyexcel::{CellValue, WriteCellData};
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct FormulaRow {
        #[excel(formula = "SUM(A1:A10)")]
        total: f64,
        raw: f64,
    }

    /// Java: `@ExcelFormula` attribute should populate the column formula field.
    #[test]
    fn t10_excel_formula07() {
        let cols = <FormulaRow as ExcelRowTrait>::schema();
        assert_eq!(cols[0].formula, Some("SUM(A1:A10)"));
        assert!(cols[1].formula.is_none());
    }

    /// Java: applying formula decoration should produce a Formula CellValue.
    #[test]
    fn t10_apply_decoration07() {
        let col = <FormulaRow as ExcelRowTrait>::schema()[0];
        let data = WriteCellData::new(CellValue::Float(0.0));
        let decorated = col.apply_decorations(data);
        match decorated.value() {
            CellValue::Formula(f) => assert_eq!(f, "SUM(A1:A10)"),
            other => panic!("expected CellValue::Formula, got {other:?}"),
        }
    }
}

mod annotation_phase1_data_validation_test {
    //! Mirrors: AnnotationDataTest#t11ExcelDataValidation07
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct ValidationRow {
        #[excel(data_validation(type = "list", operator = "between", formula1 = "A,B,C"))]
        status: String,
    }

    /// Java: `@ExcelDataValidation` attribute should populate ExcelDataValidationMeta.
    #[test]
    fn t11_excel_data_validation07() {
        let cols = <ValidationRow as ExcelRowTrait>::schema();
        let dv = cols[0]
            .data_validation
            .expect("data_validation should be set");
        assert_eq!(dv.data_type, "list");
        assert_eq!(dv.operator, "between");
        assert_eq!(dv.formula1, "A,B,C");
        assert_eq!(dv.formula2, "");
        assert!(dv.is_present());
    }
}

mod annotation_phase1_conditional_test {
    //! Mirrors: AnnotationDataTest#t12ExcelConditional07
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct ConditionalRow {
        #[excel(conditional(
            condition = "greaterThan(100)",
            font_color = "red",
            background_color = "yellow"
        ))]
        value: f64,
    }

    /// Java: `@ExcelConditional` attribute should populate conditional_format.
    #[test]
    fn t12_excel_conditional07() {
        let cols = <ConditionalRow as ExcelRowTrait>::schema();
        let cf = cols[0]
            .conditional_format
            .expect("conditional_format should be set");
        assert_eq!(cf.0, "greaterThan(100)");
        assert_eq!(cf.1, "red");
        assert_eq!(cf.2, "yellow");
    }
}

mod annotation_phase1_filter_test {
    //! Mirrors: AnnotationDataTest#t13ExcelFilter07
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct FilterRow {
        #[excel(filter)]
        name: String,
        age: u32,
    }

    /// Java: `@ExcelFilter` attribute should set auto_filter = true.
    #[test]
    fn t13_excel_filter07() {
        let cols = <FilterRow as ExcelRowTrait>::schema();
        assert!(cols[0].auto_filter);
        assert!(!cols[1].auto_filter);
    }
}

mod annotation_phase1_combined_test {
    //! Mirrors: AnnotationDataTest#t14ExcelCombined07
    //! Verifies multiple annotations can stack on the same column.
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct CombinedRow {
        #[excel(
            name = "Full Name",
            index = 0,
            format = "yyyy-MM-dd",
            column_width = 30,
            comment = "primary key",
            hyperlink = "https://docs.example.com",
            filter
        )]
        name: String,
    }

    /// Java: all derive attributes can stack on a single field.
    #[test]
    fn t14_excel_combined07() {
        let cols = <CombinedRow as ExcelRowTrait>::schema();
        let c = &cols[0];
        assert_eq!(c.name, "Full Name");
        assert_eq!(c.index, Some(0));
        assert_eq!(c.format, Some("yyyy-MM-dd"));
        assert_eq!(c.column_width, Some(30));
        assert_eq!(c.comment, Some("primary key"));
        assert_eq!(c.hyperlink, Some("https://docs.example.com"));
        assert!(c.auto_filter);
    }
}

mod annotation_phase1_metadata_test {
    //! Verifies ExcelColumn round-trips through ExcelWriteMetadata.
    use easyexcel_core::ExcelRow as ExcelRowTrait;
    use easyexcel_derive::ExcelRow;

    #[derive(Debug, ExcelRow)]
    struct MetadataRow {
        #[excel(
            image = "img.jpg",
            comment = "c",
            hyperlink = "u",
            formula = "f",
            filter
        )]
        cell: String,
    }

    /// Phase 1: ExcelColumn carries the new annotation-derived fields.
    #[test]
    fn t15_metadata_carry_through07() {
        let cols = <MetadataRow as ExcelRowTrait>::schema();
        let c = &cols[0];
        assert_eq!(c.image_path, Some("img.jpg"));
        assert_eq!(c.comment, Some("c"));
        assert_eq!(c.hyperlink, Some("u"));
        assert_eq!(c.formula, Some("f"));
        assert!(c.auto_filter);
    }
}
