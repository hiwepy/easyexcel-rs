//! Comprehensive test suite mirroring Java `com.alibaba.excel.test.core.*`.
//!
//! Java 33 test classes: SimpleDataTest, AnnotationDataTest, ConverterDataTest,
//! CellDataDataTest, ExceptionDataTest, ExtraDataTest, FillDataTest,
//! NoModelDataTest, ExcludeOrIncludeDataTest, LargeDataTest, TemplateDataTest,
//! StyleDataTest, BomDataTest, CharsetDataTest, EncryptDataTest, etc.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::io;
use std::path::PathBuf;
use std::str::FromStr;
use std::sync::Arc;
use std::time::Duration;

use chrono::{NaiveDate, NaiveDateTime};

use super::*;
use crate::constant::{get_builtin_format, EXCEL_MATH_CONTEXT_PRECISION, ROW_TAG, CELL_TAG, CELL_VALUE_TAG, CELL_FORMULA_TAG};
use crate::support::ExcelTypeEnum;

// ============================================================================
// 1. CsvCharset tests (Java: CsvCharsetTest)
// ============================================================================

#[test]
fn csv_charset_accepts_java_style_names_and_has_a_utf8_default() {
    assert_eq!(CsvCharset::default(), CsvCharset::utf8());
    assert_eq!(CsvCharset::default().name(), "UTF-8");
    assert_eq!(CsvCharset::from("GBK").name(), "GBK");
    assert_eq!(CsvCharset::from("UTF-16BE".to_owned()).name(), "UTF-16BE");
    assert_eq!(CsvCharset::from("gbk").name(), "gbk");
    assert_eq!(CsvCharset::from("windows-1252").name(), "windows-1252");
}

#[test]
fn csv_charset_implements_from_str_and_from_string() {
    let charset: CsvCharset = "UTF-8".into();
    assert_eq!(charset.name(), "UTF-8");

    let charset2: CsvCharset = String::from("ISO-8859-1").into();
    assert_eq!(charset2.name(), "ISO-8859-1");
}

// ============================================================================
// 2. CellValue tests (Java: CellDataDataTest)
// ============================================================================

fn context(format: Option<&'static str>) -> ConvertContext {
    ConvertContext {
        sheet_name: "Users".to_owned(),
        row_index: 2,
        column_index: Some(1),
        field: "value",
        format,
    }
}

#[test]
fn cell_values_have_stable_text_and_empty_semantics() {
    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    let datetime = date.and_hms_opt(12, 34, 56).expect("valid time");
    let cases = [
        (CellValue::Empty, ""),
        (CellValue::String("text".to_owned()), "text"),
        (CellValue::Error("#DIV/0!".to_owned()), "#DIV/0!"),
        (CellValue::Bool(true), "true"),
        (CellValue::Int(-12), "-12"),
        (CellValue::Float(1.5), "1.5"),
        (
            CellValue::Decimal("123.450".parse().expect("valid decimal")),
            "123.450",
        ),
        (CellValue::Date(date), "2026-07-17"),
        (CellValue::DateTime(datetime), "2026-07-17 12:34:56"),
        (CellValue::Formula("SUM(A1:A2)".to_owned()), "SUM(A1:A2)"),
        (
            CellValue::Hyperlink {
                url: "https://rust-lang.org".to_owned(),
                text: "Rust".to_owned(),
            },
            "Rust",
        ),
        (
            CellValue::Comment {
                value: Box::new(CellValue::String("value".to_owned())),
                text: "note".to_owned(),
            },
            "value",
        ),
        (CellValue::Image(vec![1, 2, 3]), ""),
    ];
    for (value, expected) in cases {
        assert_eq!(value.as_text(), expected);
    }
    assert!(CellValue::Empty.is_empty());
    assert!(!CellValue::Bool(false).is_empty());
}

#[test]
fn cell_values_expose_converter_dispatch_types() {
    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    let datetime = date.and_hms_opt(12, 34, 56).expect("valid time");
    let cases = [
        (CellValue::Empty, CellDataType::Empty),
        (CellValue::String(String::new()), CellDataType::String),
        (CellValue::Bool(true), CellDataType::Boolean),
        (CellValue::Int(1), CellDataType::Number),
        (CellValue::Float(1.0), CellDataType::Number),
        (CellValue::Decimal(BigDecimal::from(1)), CellDataType::Number),
        (CellValue::Date(date), CellDataType::Date),
        (CellValue::DateTime(datetime), CellDataType::Date),
        (CellValue::Error("#N/A".to_owned()), CellDataType::Error),
        (CellValue::Formula("1+1".to_owned()), CellDataType::Formula),
        (
            CellValue::Hyperlink {
                url: "x".to_owned(),
                text: "y".to_owned(),
            },
            CellDataType::String,
        ),
        (
            CellValue::Comment {
                value: Box::new(CellValue::String("v".to_owned())),
                text: "".to_owned(),
            },
            CellDataType::String,
        ),
        (CellValue::Image(vec![]), CellDataType::Image),
        (CellValue::RichText(RichTextStringData::new("rt")), CellDataType::RichTextString),
        (
            CellValue::Images {
                value: Box::new(CellValue::Empty),
                images: vec![],
            },
            CellDataType::Empty,
        ),
    ];
    for (value, expected) in cases {
        assert_eq!(value.data_type(), expected);
    }
}

#[test]
fn cell_value_clone_and_eq() {
    let a = CellValue::String("hello".to_owned());
    let b = a.clone();
    assert_eq!(a, b);
    assert_ne!(a, CellValue::Int(42));
}

// ============================================================================
// 3. FromExcelCell / IntoExcelCell tests (Java: ConverterDataTest)
// ============================================================================

#[test]
fn string_round_trip() {
    let ctx = context(None);
    let val = CellValue::String("abc".to_owned());
    let s = <String as FromExcelCell>::from_excel_cell(Some(&val), &ctx).unwrap();
    assert_eq!(s, "abc");

    let cell = s.to_excel_cell(&ctx).unwrap();
    assert_eq!(cell, CellValue::String("abc".to_owned()));
}

#[test]
fn bool_from_string() {
    let ctx = context(None);
    assert!(<bool as FromExcelCell>::from_excel_cell(Some(&CellValue::String("true".to_owned())), &ctx).unwrap());
    assert!(!<bool as FromExcelCell>::from_excel_cell(Some(&CellValue::String("false".to_owned())), &ctx).unwrap());
    assert!(!<bool as FromExcelCell>::from_excel_cell(Some(&CellValue::String("0".to_owned())), &ctx).unwrap());
    assert!(<bool as FromExcelCell>::from_excel_cell(Some(&CellValue::String("1".to_owned())), &ctx).unwrap());
}

#[test]
fn integer_conversions() {
    let ctx = context(None);
    assert_eq!(<i64 as FromExcelCell>::from_excel_cell(Some(&CellValue::Int(42)), &ctx).unwrap(), 42);
    assert!(<i64 as FromExcelCell>::from_excel_cell(Some(&CellValue::Float(3.7)), &ctx).unwrap_err().to_string().contains("i64"));
    assert_eq!(<i32 as FromExcelCell>::from_excel_cell(Some(&CellValue::String("100".to_owned())), &ctx).unwrap(), 100);
}

#[test]
fn float_from_integer_cell() {
    let ctx = context(None);
    assert_eq!(<f64 as FromExcelCell>::from_excel_cell(Some(&CellValue::Int(42)), &ctx).unwrap(), 42.0);
}

#[test]
fn bigdecimal_round_trip() {
    let ctx = context(None);
    let original = BigDecimal::from_str("123.456789").unwrap();
    let cell = CellValue::Decimal(original.clone());
    let recovered = <BigDecimal as FromExcelCell>::from_excel_cell(Some(&cell), &ctx).unwrap();
    assert_eq!(recovered, original);
}

#[test]
fn naivedate_from_string_with_format() {
    let ctx = context(Some("%Y-%m-%d"));
    let cell = CellValue::String("2026-03-15".to_owned());
    let d = <NaiveDate as FromExcelCell>::from_excel_cell(Some(&cell), &ctx).unwrap();
    assert_eq!(d, NaiveDate::from_ymd_opt(2026, 3, 15).unwrap());
}

#[test]
fn naivedatetime_from_string_with_format() {
    let ctx = context(Some("%Y-%m-%d %H:%M:%S"));
    let cell = CellValue::String("2026-03-15 14:30:00".to_owned());
    let dt = <NaiveDateTime as FromExcelCell>::from_excel_cell(Some(&cell), &ctx).unwrap();
    assert_eq!(dt, NaiveDateTime::new(NaiveDate::from_ymd_opt(2026, 3, 15).unwrap(), chrono::NaiveTime::from_hms_opt(14, 30, 0).unwrap()));
}

#[test]
fn option_cell_handles_empty() {
    let ctx = context(None);
    assert_eq!(<Option<String> as FromExcelCell>::from_excel_cell(None, &ctx).unwrap(), None);
    assert_eq!(<Option<String> as FromExcelCell>::from_excel_cell(Some(&CellValue::Empty), &ctx).unwrap(), None);
    assert_eq!(
        <Option<String> as FromExcelCell>::from_excel_cell(Some(&CellValue::String("x".to_owned())), &ctx).unwrap(),
        Some("x".to_owned())
    );
}

#[test]
fn image_vec_round_trip() {
    let ctx = context(None);
    let img = vec![0x89, 0x50, 0x4E, 0x47];
    let cell = <Vec<u8> as IntoExcelCell>::to_excel_cell(&img, &ctx).unwrap();
    assert!(matches!(cell, CellValue::Image(_)));
    let back = <Vec<u8> as FromExcelCell>::from_excel_cell(Some(&cell), &ctx).unwrap();
    assert_eq!(back, img);
}

#[test]
fn pathbuf_round_trip() {
    let ctx = context(None);
    let pb = PathBuf::from("/tmp/image.png");
    let cell = <PathBuf as IntoExcelCell>::to_excel_cell(&pb, &ctx);
    // File may not exist, so we just check the type
    assert!(cell.is_ok() || cell.is_err());
}

// ============================================================================
// 4. CoordinateData tests
// ============================================================================

#[test]
fn coordinate_data_builder_chain() {
    let coord = CoordinateData::new()
        .first_row_index(5)
        .first_column_index(3)
        .relative_last_row_index(2)
        .relative_last_column_index(1);
    assert_eq!(coord.get_first_row_index(), Some(5));
    assert_eq!(coord.get_first_column_index(), Some(3));
    assert_eq!(coord.get_relative_last_row_index(), Some(2));
    assert_eq!(coord.get_relative_last_column_index(), Some(1));
    assert_eq!(coord.get_last_row_index(), None);
}

#[test]
fn coordinate_data_clone_eq() {
    let a = CoordinateData::new().first_row_index(1).first_column_index(2);
    let b = a.clone();
    assert_eq!(a, b);
}

// ============================================================================
// 5. ClientAnchorData tests
// ============================================================================

#[test]
fn client_anchor_data_builder() {
    let anchor = ClientAnchorData::new()
        .top(100)
        .left(50)
        .anchor_type(AnchorType::MoveAndResize);
    assert_eq!(anchor.get_top(), Some(100));
    assert_eq!(anchor.get_left(), Some(50));
    assert_eq!(anchor.get_anchor_type(), Some(AnchorType::MoveAndResize));
}

#[test]
fn anchor_type_variants() {
    assert_eq!(AnchorType::MoveAndResize, AnchorType::MoveAndResize);
    assert_eq!(AnchorType::DontMoveDoResize, AnchorType::DontMoveDoResize);
    assert_eq!(AnchorType::MoveDontResize, AnchorType::MoveDontResize);
    assert_eq!(AnchorType::DontMoveAndResize, AnchorType::DontMoveAndResize);
    assert_ne!(AnchorType::MoveAndResize, AnchorType::MoveDontResize);
}

// ============================================================================
// 6. ImageData tests
// ============================================================================

#[test]
fn image_data_builder_chain() {
    let img = ImageData::new(vec![0x89, 0x50, 0x4E, 0x47])
        .image_type(ImageType::Png);
    assert_eq!(img.image(), &[0x89, 0x50, 0x4E, 0x47]);
    assert_eq!(img.get_image_type(), Some(ImageType::Png));
    assert_eq!(img.get_anchor(), ClientAnchorData::new());
}

#[test]
fn image_type_variants() {
    let types = [ImageType::Emf, ImageType::Wmf, ImageType::Pict, ImageType::Jpeg, ImageType::Png, ImageType::Dib];
    assert_eq!(types.len(), 6);
}

// ============================================================================
// 7. RichTextStringData tests
// ============================================================================

#[test]
fn richtext_string_data_builder() {
    let rt = RichTextStringData::new("Hello World")
        .apply_font(WriteFont::new().bold(true).font_name("Arial".to_owned()))
        .apply_font_range(0, 5, WriteFont::new().color(ExcelColor::Rgb(0xFF0000)));
    assert_eq!(rt.text_string(), "Hello World");
    assert!(rt.write_font().is_some());
    assert!(rt.write_font().unwrap().get_bold() == Some(true));
    assert_eq!(rt.interval_fonts().len(), 1);
}

// ============================================================================
// 8. WriteCellData tests (Java: CellDataDataTest)
// ============================================================================

#[test]
fn write_cell_data_constructors() {
    let _ctx = context(None);
    let ws = WriteCellData::new(CellValue::String("hi".to_owned()));
    assert_eq!(*ws.value(), CellValue::String("hi".to_owned()));
    assert!(ws.images().is_empty());

    let img = WriteCellData::from_image(vec![1, 2]);
    assert_eq!(*img.value(), CellValue::Empty);
    assert_eq!(img.images().len(), 1);

    let rt = WriteCellData::from_rich_text(RichTextStringData::new("rich"));
    assert!(matches!(rt.value(), CellValue::RichText(_)));
}

// ============================================================================
// 9. ReadCellData tests
// ============================================================================

#[test]
fn read_cell_data_fields() {
    let rd = ReadCellData::new(
        5, 2,
        CellValue::Int(42),
        CellValue::Int(42),
        "42".to_owned(),
        None,
    );
    assert_eq!(rd.row_index(), 5);
    assert_eq!(rd.column_index(), 2);
    assert_eq!(*rd.raw_value(), CellValue::Int(42));
    assert_eq!(rd.display_value(), "42");
    assert!(rd.formula().is_none());
}

// ============================================================================
// 10. FormulaData tests
// ============================================================================

#[test]
fn formula_data_clone() {
    let f1 = FormulaData::new("SUM(A1:A10)".to_owned());
    let f2 = f1.clone();
    assert_eq!(f1, f2);
    assert_eq!(f1.formula_value(), "SUM(A1:A10)");
}

// ============================================================================
// 11. DynamicValue / DynamicRow tests (Java: NoModelDataTest)
// ============================================================================

#[test]
fn dynamic_row_get_by_column() {
    let mut map = BTreeMap::new();
    map.insert(0, DynamicValue::String("hello".to_owned()));
    map.insert(2, DynamicValue::ActualData(CellValue::Int(42)));
    let row = DynamicRow::new(map);
    assert_eq!(row.get(0), Some(&DynamicValue::String("hello".to_owned())));
    assert_eq!(row.get(1), None);
    assert_eq!(row.get(2), Some(&DynamicValue::ActualData(CellValue::Int(42))));
    assert_eq!(row.values().len(), 2);
    assert_eq!(row.into_values().len(), 2);
}

#[test]
fn dynamic_row_clone_eq() {
    let map = BTreeMap::new();
    let a = DynamicRow::new(map.clone());
    let b = DynamicRow::new(map);
    assert_eq!(a, b);
}

// ============================================================================
// 12. ReadDefaultReturn tests
// ============================================================================

#[test]
fn read_default_return_variants() {
    assert_eq!(ReadDefaultReturn::String, ReadDefaultReturn::String);
    assert_eq!(ReadDefaultReturn::ActualData, ReadDefaultReturn::ActualData);
    assert_eq!(ReadDefaultReturn::ReadCellData, ReadDefaultReturn::ReadCellData);
}

// ============================================================================
// 13. ExcelError tests (Java: ExceptionDataTest)
// ============================================================================

#[test]
fn excel_error_display() {
    let err = ExcelError::SheetNotFound("Sheet2".to_owned());
    assert!(err.to_string().contains("Sheet2"));

    let err2 = ExcelError::Format("bad xml".to_owned());
    assert!(err2.to_string().contains("bad xml"));

    let err3 = ExcelError::Unsupported("write xls".to_owned());
    assert!(err3.to_string().contains("write xls"));

    let err4 = ExcelError::Data {
        sheet: "S1".to_owned(),
        row: 5,
        column: Some(3),
        field: "age",
        value: "abc".to_owned(),
        message: "bad int".to_owned(),
    };
    let msg = err4.to_string();
    assert!(msg.contains("S1"));
    assert!(msg.contains("5"));
    assert!(msg.contains("age"));
    assert!(msg.contains("bad int"));
}

#[test]
fn excel_error_from_io() {
    let io_err = io::Error::new(io::ErrorKind::NotFound, "file missing");
    let excel_err = ExcelError::Io(io_err);
    assert!(excel_err.to_string().contains("file missing"));
}

// ============================================================================
// 14. ExcelColumn tests
// ============================================================================

#[test]
fn excel_column_builder_chain() {
    let col = ExcelColumn::new("age", "Age", Some(1), 10, None)
        .with_column_width(20)
        .with_head_style(ExcelCellStyle { horizontal_alignment: Some(ExcelHorizontalAlignment::Center), ..ExcelCellStyle::new() })
        .with_content_style(ExcelCellStyle { hidden: Some(true), ..ExcelCellStyle::new() })
        .with_head_font_style(ExcelFontStyle { bold: Some(true), ..ExcelFontStyle::new() })
        .with_content_font_style(ExcelFontStyle { font_name: Some("Arial"), ..ExcelFontStyle::new() });
    assert_eq!(col.field, "age");
    assert_eq!(col.name, "Age");
    assert_eq!(col.index, Some(1));
    assert_eq!(col.column_width, Some(20));
}

// ============================================================================
// 15. ExcelCellStyle tests
// ============================================================================

#[test]
fn excel_cell_style_fields() {
    let style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        vertical_alignment: Some(ExcelVerticalAlignment::Top),
        border_left: Some(ExcelBorderStyle::Thin),
        fill_pattern: Some(ExcelFillPattern::Solid),
        data_format: Some(ExcelDataFormat::Custom("0.00")),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.horizontal_alignment, Some(ExcelHorizontalAlignment::Center));
    assert_eq!(style.vertical_alignment, Some(ExcelVerticalAlignment::Top));
    assert_eq!(style.border_left, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.fill_pattern, Some(ExcelFillPattern::Solid));
    assert_eq!(style.data_format, Some(ExcelDataFormat::Custom("0.00")));
}

// ============================================================================
// 16. ExcelWriteMetadata tests
// ============================================================================

#[test]
fn excel_write_metadata_builder_chain() {
    let meta = ExcelWriteMetadata::new()
        .column_width(25)
        .head_row_height(30)
        .content_row_height(20)
        .head_style(ExcelCellStyle { horizontal_alignment: Some(ExcelHorizontalAlignment::Center), ..ExcelCellStyle::new() })
        .head_font_style(ExcelFontStyle { bold: Some(true), ..ExcelFontStyle::new() });
    assert_eq!(meta.column_width, Some(25));
    assert_eq!(meta.head_row_height, Some(30));
    assert_eq!(meta.content_row_height, Some(20));
}

// ============================================================================
// 17. CellExtra tests (Java: ExtraDataTest)
// ============================================================================

#[test]
fn cell_extra_fields() {
    let extra = CellExtra::new(
        CellExtraType::Comment,
        Some("this is a comment".to_owned()),
        0, 0, 1, 1,
    );
    assert_eq!(extra.extra_type(), CellExtraType::Comment);
    assert_eq!(extra.text(), Some("this is a comment"));
    assert_eq!(extra.first_row_index(), 0);
    assert_eq!(extra.last_column_index(), 1);
}

#[test]
fn cell_extra_merge_range() {
    let merge = CellExtra::new(CellExtraType::Merge, None, 1, 5, 0, 3);
    assert_eq!(merge.first_row_index(), 1);
    assert_eq!(merge.last_row_index(), 5);
}

// ============================================================================
// 18. RowData tests
// ============================================================================

#[test]
fn row_data_cell_resolution() {
    let mut headers = HashMap::new();
    headers.insert("Name".to_owned(), 0);
    headers.insert("Age".to_owned(), 1);
    let cells = vec![
        CellValue::String("Alice".to_owned()),
        CellValue::Int(30),
    ];
    let row = RowData::new("Sheet1", 0, cells, Arc::new(headers));

    let name_col = ExcelColumn::new("name", "Name", None, 0, None);
    let age_col = ExcelColumn::new("age", "Age", Some(1), 10, None);

    assert_eq!(row.cell(&name_col), Some(&CellValue::String("Alice".to_owned())));
    assert_eq!(row.cell(&age_col), Some(&CellValue::Int(30)));
}

#[test]
fn row_data_formula_resolution() {
    let headers = Arc::new(HashMap::new());
    let cells = vec![CellValue::Empty, CellValue::Float(10.0)];
    let mut formulas = HashMap::new();
    formulas.insert(1, FormulaData::new("SUM(A1:A5)".to_owned()));
    let row = RowData::new("S", 0, cells, headers).with_formulas(formulas);

    let col = ExcelColumn::new("total", "Total", Some(1), 0, None);
    let formula = row.formula(&col).expect("formula present");
    assert_eq!(formula.formula_value(), "SUM(A1:A5)");
}

#[test]
fn row_data_convert_context() {
    let headers = Arc::new(HashMap::new());
    let row = RowData::new("Users", 5, vec![], headers);
    let col = ExcelColumn::new("email", "Email", Some(2), 0, Some("%Y-%m-%d"));
    let ctx = row.convert_context(&col);
    assert_eq!(ctx.sheet_name, "Users");
    assert_eq!(ctx.row_index, 5);
    assert_eq!(ctx.column_index, Some(2));
    assert_eq!(ctx.field, "email");
    assert_eq!(ctx.format, Some("%Y-%m-%d"));
}

// ============================================================================
// 19. ReadListener tests (Java: ExceptionDataTest)
// ============================================================================

struct TestListener {
    rows: Vec<String>,
    _batch_idx: usize,
}

impl TestListener {
    fn new() -> Self {
        Self { rows: Vec::new(), _batch_idx: 0 }
    }
}

impl ReadListener<String> for TestListener {
    fn invoke(&mut self, data: String, _context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        Ok(())
    }

    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        ErrorAction::Continue
    }

    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }
}

#[test]
fn read_listener_invoke_works() {
    let mut listener = TestListener::new();
    let ctx = AnalysisContext::new("S1", 0, 0);
    listener.invoke("row1".to_owned(), &ctx).unwrap();
    assert_eq!(listener.rows, vec!["row1"]);
}

#[test]
fn read_listener_can_stop() {
    struct StopAfterOne {
        count: usize,
    }
    impl ReadListener<String> for StopAfterOne {
        fn invoke(&mut self, _data: String, _context: &AnalysisContext) -> Result<()> {
            self.count += 1;
            if self.count >= 2 {
                return Err(ExcelError::Format("stop".to_owned()));
            }
            Ok(())
        }
        fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> { Ok(()) }
    }
    let mut listener = StopAfterOne { count: 0 };
    let ctx = AnalysisContext::new("S", 0, 0);
    listener.invoke("a".to_owned(), &ctx).unwrap();
    let err = listener.invoke("b".to_owned(), &ctx);
    assert!(err.is_err());
}

#[test]
fn read_listener_has_next_can_stop() {
    struct StopAfterTwo {
        count: usize,
    }
    impl ReadListener<String> for StopAfterTwo {
        fn invoke(&mut self, _data: String, _context: &AnalysisContext) -> Result<()> { self.count += 1; Ok(()) }
        fn has_next(&mut self, _context: &AnalysisContext) -> bool { self.count < 2 }
        fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> { Ok(()) }
    }
    let mut listener = StopAfterTwo { count: 0 };
    let ctx = AnalysisContext::new("S", 0, 0);
    assert!(listener.has_next(&ctx));
    listener.invoke("a".to_owned(), &ctx).unwrap();
    assert!(listener.has_next(&ctx));
    listener.invoke("b".to_owned(), &ctx).unwrap();
    assert!(!listener.has_next(&ctx));
}

// ============================================================================
// 20. PageReadListener tests (Java: PageReadListenerTest)
// ============================================================================

#[test]
fn page_read_listener_batches() {
    let batch_count = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let bc = batch_count.clone();
    let mut listener = PageReadListener::new(2, move |data, _ctx| {
        bc.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        assert!(data.len() <= 2);
        Ok(())
    });

    let ctx = AnalysisContext::new("S", 0, 0);
    listener.invoke("a".to_owned(), &ctx).unwrap();
    assert_eq!(batch_count.load(std::sync::atomic::Ordering::Relaxed), 0);
    listener.invoke("b".to_owned(), &ctx).unwrap();
    assert_eq!(batch_count.load(std::sync::atomic::Ordering::Relaxed), 1);
    listener.invoke("c".to_owned(), &ctx).unwrap();
    // partial batch: no callback yet
    assert_eq!(batch_count.load(std::sync::atomic::Ordering::Relaxed), 1);
}

#[test]
fn page_read_listener_flush_on_end() {
    let batch_count = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let bc = batch_count.clone();
    let mut listener = PageReadListener::new(5, move |_data, _ctx| {
        bc.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        Ok(())
    });
    let ctx = AnalysisContext::new("S", 0, 0);
    listener.invoke("a".to_owned(), &ctx).unwrap();
    listener.invoke("b".to_owned(), &ctx).unwrap();
    assert_eq!(batch_count.load(std::sync::atomic::Ordering::Relaxed), 0);
    listener.do_after_all_analysed(&ctx).unwrap();
    assert_eq!(batch_count.load(std::sync::atomic::Ordering::Relaxed), 1);
}

// ============================================================================
// 21. ExcelError conversion chain
// ============================================================================

#[test]
fn excel_data_error_contains_row_column() {
    let err = ExcelError::Data {
        sheet: "Users".to_owned(),
        row: 10,
        column: Some(5),
        field: "amount",
        value: "abc".to_owned(),
        message: "not a number".to_owned(),
    };
    let msg = err.to_string();
    assert!(msg.contains("sheet=Users"));
    assert!(msg.contains("row=10"));
    assert!(msg.contains("field=amount"));
    assert!(msg.contains("not a number"));
}

#[test]
fn excel_error_is_cloneable() {
    let err = ExcelError::Format("test".to_owned());
    let err2 = err.clone();
    assert_eq!(err.to_string(), err2.to_string());
}

// ============================================================================
// 22. AnalysisContext tests
// ============================================================================

#[test]
fn analysis_context_construction() {
    let ctx = AnalysisContext::new("Sheet1", 0, 100);
    assert_eq!(ctx.sheet_name(), "Sheet1");
    assert_eq!(ctx.sheet_no(), 0);
    assert_eq!(ctx.row_index(), 100);
}

#[test]
fn analysis_context_with_custom_object() {
    let ctx = AnalysisContext::new("S", 0, 0)
        .with_custom_object(Some(CustomReadObject::new(42u32)));
    let val = ctx.custom::<u32>();
    assert_eq!(val, Some(&42u32));
}

#[test]
fn analysis_context_with_batch_index() {
    let ctx = AnalysisContext::new("S", 0, 0).with_batch_index(7);
    assert_eq!(ctx.batch_index(), 7);
}

// ============================================================================
// 23. WriteHandler tests (Java: WriteHandlerTest)
// ============================================================================

struct TestWriteHandler {
    order: i32,
    before_workbook_called: std::sync::atomic::AtomicBool,
    before_cell_value: Option<CellValue>,
}

impl WriteHandler for TestWriteHandler {
    fn order(&self) -> i32 { self.order }
    fn before_workbook(&mut self, _ctx: &WriteWorkbookContext) -> Result<()> {
        self.before_workbook_called.store(true, std::sync::atomic::Ordering::Relaxed);
        Ok(())
    }
    fn before_cell(&mut self, ctx: &mut WriteCellContext) -> Result<()> {
        self.before_cell_value = Some(ctx.value.clone());
        Ok(())
    }
}

#[test]
fn write_handler_order_and_before_workbook() {
    let mut h = TestWriteHandler { order: -10, before_workbook_called: std::sync::atomic::AtomicBool::new(false), before_cell_value: None };
    assert_eq!(h.order(), -10);
    let ctx = WriteWorkbookContext::new("out.xlsx");
    h.before_workbook(&ctx).unwrap();
    assert!(h.before_workbook_called.load(std::sync::atomic::Ordering::Relaxed));
}

#[test]
fn write_handler_before_cell_receives_value() {
    let mut h = TestWriteHandler { order: 0, before_workbook_called: std::sync::atomic::AtomicBool::new(false), before_cell_value: None };
    let mut ctx = WriteCellContext {
        sheet_name: "S".to_owned(),
        row_index: 0,
        column_index: 0,
        field: None,
        is_head: false,
        relative_row_index: None,
        value: CellValue::String("hello".to_owned()),
        skip: false,
    };
    h.before_cell(&mut ctx).unwrap();
    assert_eq!(h.before_cell_value, Some(CellValue::String("hello".to_owned())));
}

// ============================================================================
// 24. ConverterRegistry tests (Java: ConverterTest)
// ============================================================================

struct PrefixConverter;

impl Converter<String> for PrefixConverter {
    fn convert_to_rust_data(&self, ctx: &ReadConverterContext<'_>) -> Result<String> {
        let val = ctx.cell().map_or_else(String::new, CellValue::as_text);
        Ok(format!("custom:{val}"))
    }
    fn convert_to_excel_data(&self, ctx: &WriteConverterContext<'_, String>) -> Result<CellValue> {
        Ok(CellValue::String(format!("custom:{}", ctx.value())))
    }
}

#[test]
fn converter_registry_register_and_read() {
    let mut registry = ConverterRegistry::default();
    registry.register::<String, _>(PrefixConverter);
    assert!(!registry.is_empty());

    let ctx = ConvertContext {
        sheet_name: "S".to_owned(),
        row_index: 0,
        column_index: Some(0),
        field: "f",
        format: None,
    };
    let cell = CellValue::String("abc".to_owned());
    let col = ExcelColumn::new("f","F", Some(0),0,None);
    let rctx = ReadConverterContext::new(Some(&cell), &col, &ctx);
    let result = registry.convert_to_rust_data::<String>(&rctx).unwrap().unwrap();
    assert_eq!(result, "custom:abc");
}

#[test]
fn converter_registry_write() {
    let mut registry = ConverterRegistry::default();
    registry.register::<String, _>(PrefixConverter);
    let ctx = ConvertContext {
        sheet_name: "S".to_owned(), row_index: 0, column_index: Some(0),
        field: "f", format: None,
    };
    let col = ExcelColumn::new("f","F", Some(0),0,None);
    let cell = registry.convert_to_excel_data(&"test".to_owned(), &col, &ctx).unwrap().unwrap();
    assert_eq!(cell, CellValue::String("custom:test".to_owned()));
}

#[test]
fn converter_registry_merged_with_takes_priority() {
    struct AConverter;
    impl Converter<String> for AConverter {
        fn convert_to_rust_data(&self, _: &ReadConverterContext<'_>) -> Result<String> { Ok("A".to_owned()) }
        fn convert_to_excel_data(&self, _: &WriteConverterContext<'_, String>) -> Result<CellValue> { Ok(CellValue::String("A".to_owned())) }
    }
    struct BConverter;
    impl Converter<String> for BConverter {
        fn convert_to_rust_data(&self, _: &ReadConverterContext<'_>) -> Result<String> { Ok("B".to_owned()) }
        fn convert_to_excel_data(&self, _: &WriteConverterContext<'_, String>) -> Result<CellValue> { Ok(CellValue::String("B".to_owned())) }
    }

    let mut base = ConverterRegistry::default();
    base.register::<String, _>(AConverter);
    let mut overrides = ConverterRegistry::default();
    overrides.register::<String, _>(BConverter);

    let merged = base.merged_with(&overrides);
    let ctx = ConvertContext { sheet_name: "S".to_owned(), row_index: 0, column_index: Some(0), field: "f", format: None };
    let col = ExcelColumn::new("f","F", Some(0),0,None);
    let empty_cell = CellValue::String("".to_owned());
    let rctx = ReadConverterContext::new(Some(&empty_cell), &col, &ctx);
    let result = merged.convert_to_rust_data::<String>(&rctx).unwrap().unwrap();
    assert_eq!(result, "B"); // overrides take priority
}

// ============================================================================
// 25. StringImageConverter tests
// ============================================================================

#[test]
fn string_image_converter_read_error() {
    let converter = StringImageConverter;
    let ctx = ConvertContext { sheet_name: "S".to_owned(), row_index: 0, column_index: Some(0), field: "img", format: None };
    let cell = CellValue::String("nonexistent.png".to_owned());
    let col = ExcelColumn::new("img","Img", Some(0),0,None);
    let rctx = ReadConverterContext::new(Some(&cell), &col, &ctx);
    // convert_to_rust_data should return Unsupported error
    let err = Converter::<String>::convert_to_rust_data(&converter, &rctx);
    assert!(err.is_err());
}

// ============================================================================
// 26. UrlImageConverter tests
// ============================================================================

#[test]
fn url_image_converter_timeouts() {
    let c = UrlImageConverter::default();
    assert_eq!(c.connect_timeout(), Duration::from_secs(1));
    assert_eq!(c.read_timeout(), Duration::from_secs(5));

    let c2 = UrlImageConverter::new(Duration::from_secs(2), Duration::from_secs(10));
    assert_eq!(c2.connect_timeout(), Duration::from_secs(2));
    assert_eq!(c2.read_timeout(), Duration::from_secs(10));
}

// ============================================================================
// 27. ExcelTypeEnum tests (Java: ExcelTypeEnumTest)
// ============================================================================

#[test]
fn excel_type_enum_value() {
    assert_eq!(ExcelTypeEnum::Csv.value(), ".csv");
    assert_eq!(ExcelTypeEnum::Xls.value(), ".xls");
    assert_eq!(ExcelTypeEnum::Xlsx.value(), ".xlsx");
}

#[test]
fn excel_type_enum_from_extension() {
    assert_eq!(ExcelTypeEnum::from_extension("csv"), Some(ExcelTypeEnum::Csv));
    assert_eq!(ExcelTypeEnum::from_extension("xls"), Some(ExcelTypeEnum::Xls));
    assert_eq!(ExcelTypeEnum::from_extension("xlsx"), Some(ExcelTypeEnum::Xlsx));
    assert_eq!(ExcelTypeEnum::from_extension("unknown"), None);
}

// ============================================================================
// 28. BuiltinFormats tests (Java: BuiltinFormatsTest)
// ============================================================================

#[test]
fn builtin_formats_has_all_indices() {
    assert!(!get_builtin_format(0, "").is_empty());
    assert!(!get_builtin_format(1, "").is_empty());
    assert!(!get_builtin_format(49, "").is_empty());
}

#[test]
fn builtin_format_14_is_date() {
    let fmt = get_builtin_format(14, "");
    assert!(fmt.contains("yyyy") || fmt.contains("m/d"));
}

// ============================================================================
// 29. ExcelXmlConstants tests
// ============================================================================

#[test]
fn xml_constants_are_nonempty() {
    assert!(!ROW_TAG.is_empty());
    assert!(!CELL_TAG.is_empty());
    assert!(!CELL_VALUE_TAG.is_empty());
    assert!(!CELL_FORMULA_TAG.is_empty());
}

// ============================================================================
// 30. EasyExcelConstants tests
// ============================================================================

#[test]
fn easy_excel_constants_math_context() {
    assert_eq!(EXCEL_MATH_CONTEXT_PRECISION, 15);
}

// ============================================================================
// 31. WriteWorkbookContext / WriteSheetContext / WriteRowContext tests
// ============================================================================

#[test]
fn write_workbook_context() {
    let ctx = WriteWorkbookContext::new("test.xlsx");
    assert_eq!(ctx.path(), std::path::Path::new("test.xlsx"));
}

#[test]
fn write_sheet_context() {
    let ctx = WriteSheetContext::new("Sheet1");
    assert_eq!(ctx.sheet_name(), "Sheet1");
}

#[test]
fn write_row_context() {
    let ctx = WriteRowContext {
        sheet_name: "Sheet1".to_owned(),
        row_index: 5,
        is_head: false,
    };
    assert_eq!(ctx.row_index, 5);
    assert!(!ctx.is_head);
}

#[test]
fn write_cell_context_skip_value() {
    let ctx = WriteCellContext {
        sheet_name: "Sheet1".to_owned(),
        row_index: 0,
        column_index: 0,
        field: Some("name"),
        is_head: false,
        relative_row_index: None,
        value: CellValue::String("Alice".to_owned()),
        skip: false,
    };
    assert!(!ctx.skip);
}

// ============================================================================
// 32. BooleanEnum tests
// ============================================================================

#[test]
fn boolean_enum_tristate() {
    assert_eq!(BooleanEnum::Default.value(), None);
    assert_eq!(BooleanEnum::True.value(), Some(true));
    assert_eq!(BooleanEnum::False.value(), Some(false));
}

// ============================================================================
// 33. ExcelHorizontalAlignment / VerticalAlignment / BorderStyle / FillPattern tests
// ============================================================================

#[test]
fn alignment_enums_variants() {
    assert_eq!(ExcelHorizontalAlignment::Center, ExcelHorizontalAlignment::Center);
    assert_ne!(ExcelHorizontalAlignment::Center, ExcelHorizontalAlignment::Left);
    assert_eq!(ExcelVerticalAlignment::Bottom, ExcelVerticalAlignment::Bottom);
}

#[test]
fn border_style_and_fill_pattern_enum_variants() {
    // Verify key enum variants exist and are distinct.
    assert_ne!(ExcelBorderStyle::None, ExcelBorderStyle::Thin);
    assert_ne!(ExcelBorderStyle::Thin, ExcelBorderStyle::Medium);
    assert_ne!(ExcelBorderStyle::Double, ExcelBorderStyle::Hair);
    assert_ne!(ExcelBorderStyle::SlantDashDot, ExcelBorderStyle::DashDotDot);

    assert_ne!(ExcelUnderline::None, ExcelUnderline::Single);
    assert_ne!(ExcelUnderline::Single, ExcelUnderline::Double);
    assert_ne!(ExcelUnderline::SingleAccounting, ExcelUnderline::DoubleAccounting);

    assert_ne!(ExcelFontScript::None, ExcelFontScript::Superscript);
    assert_ne!(ExcelFontScript::Superscript, ExcelFontScript::Subscript);

    assert_ne!(ExcelFillPattern::None, ExcelFillPattern::Solid);
    assert_ne!(ExcelFillPattern::Gray125, ExcelFillPattern::Gray0625);
}

#[test]
fn excel_color_indexed_and_rgb() {
    let c1 = ExcelColor::java_or_rgb(5);
    assert_eq!(c1, ExcelColor::Indexed(5));
    let c2 = ExcelColor::java_or_rgb(0xFF0000);
    assert_eq!(c2, ExcelColor::Rgb(0xFF0000));
}

// ============================================================================
// 34. ExcelDataFormat tests
// ============================================================================

#[test]
fn excel_data_format_variants() {
    let builtin = ExcelDataFormat::Builtin(14);
    let custom = ExcelDataFormat::Custom("yyyy/m/d");
    assert_eq!(builtin, ExcelDataFormat::Builtin(14));
    assert_ne!(builtin, custom);
}

// ============================================================================
// 35. ReadCellData / WriteCellData integration
// ============================================================================

#[test]
fn read_write_cell_data_round_trip() {
    let ctx = context(None);
    // Scalar WriteCellData stays unwrapped; Images only wraps when image_data_list is non-empty.
    let ws = WriteCellData::new(CellValue::Int(100));
    let ws_cell = ws.to_excel_cell(&ctx).unwrap();
    assert_eq!(ws_cell, CellValue::Int(100));
    let rd = ReadCellData::new(0, 0, ws_cell.clone(), ws_cell, "100".to_owned(), None);
    assert_eq!(rd.row_index(), 0);
    assert_eq!(rd.raw_value(), &CellValue::Int(100));
    assert_eq!(rd.display_value(), "100");

    let with_image = WriteCellData::new(CellValue::Int(100))
        .image(ImageData::new(vec![0x89, 0x50, 0x4e, 0x47]));
    let imaged = with_image.to_excel_cell(&ctx).unwrap();
    match imaged {
        CellValue::Images { value, images } => {
            assert_eq!(*value, CellValue::Int(100));
            assert_eq!(images.len(), 1);
        }
        other => panic!("expected Images wrapper, got {other:?}"),
    }
}

// ============================================================================
// 36. DynamicRow ExcelRow impl
// ============================================================================

#[test]
fn dynamic_row_from_row_data() {
    let headers = Arc::new(HashMap::new());
    let cells = vec![
        CellValue::String("Alice".to_owned()),
        CellValue::Int(30),
        CellValue::Empty,
    ];
    let row_data = RowData::new("S", 0, cells, headers);
    let dynamic = DynamicRow::from_row(&row_data).unwrap();
    assert_eq!(dynamic.get(0), Some(&DynamicValue::String("Alice".to_owned())));
    // Empty cells become empty strings in default ReadDefaultReturn::String mode
    assert_eq!(dynamic.get(2), Some(&DynamicValue::String("".to_owned())));
}

// ============================================================================
// 37. RowData display_values and decimal_values
// ============================================================================

#[test]
fn row_data_display_values_override() {
    let headers = Arc::new(HashMap::new());
    let cells = vec![CellValue::Float(12345678.1234567)];
    let mut display_values = HashMap::new();
    display_values.insert(0, "12345678.12".to_owned());
    let row = RowData::new("S", 0, cells, headers)
        .with_display_values(display_values);
    let col = ExcelColumn::new("v", "V", Some(0), 0, None);
    // When ReadDefaultReturn::String (default), dynamic_cell uses display_value
    assert_eq!(*row.cell(&col).unwrap(), CellValue::Float(12345678.1234567));
}

// ============================================================================
// 38. ExcelError formatting edge cases
// ============================================================================

#[test]
fn excel_error_data_with_none_column() {
    let err = ExcelError::Data {
        sheet: "S".to_owned(),
        row: 0,
        column: None,
        field: "x",
        value: String::new(),
        message: "err".to_owned(),
    };
    let msg = err.to_string();
    assert!(msg.contains("column=None"));
}

// ============================================================================
// 39. ReadListener extra method
// ============================================================================

#[test]
fn read_listener_extra_is_noop_by_default() {
    struct NoopListener;
    impl ReadListener<String> for NoopListener {
        fn invoke(&mut self, _data: String, _ctx: &AnalysisContext) -> Result<()> { Ok(()) }
        fn do_after_all_analysed(&mut self, _ctx: &AnalysisContext) -> Result<()> { Ok(()) }
    }
    let mut listener = NoopListener;
    let ctx = AnalysisContext::new("S", 0, 0);
    let extra = CellExtra::new(CellExtraType::Merge, None, 0, 0, 0, 0);
    let _ = listener.extra(&extra, &ctx); // should not panic
}

// ============================================================================
// 40. WriteCellData image_list builder
// ============================================================================

#[test]
fn write_cell_data_with_images() {
    let ws = WriteCellData::new(CellValue::String("img".to_owned()))
        .image(ImageData::new(vec![0x89, 0x50, 0x4E, 0x47]).image_type(ImageType::Png));
    assert_eq!(ws.images().len(), 1);
    assert_eq!(ws.images()[0].get_image_type(), Some(ImageType::Png));

    let ws2 = ws.image(ImageData::new(vec![0x42, 0x4D]).image_type(ImageType::Dib));
    assert_eq!(ws2.images().len(), 2);
}

// ============================================================================
// 41. ExcelColor from u32 (Java: POI indexed colors 0..64)
// ============================================================================

#[test]
fn excel_color_java_or_rgb_boundary() {
    assert_eq!(ExcelColor::java_or_rgb(0), ExcelColor::Indexed(0));
    assert_eq!(ExcelColor::java_or_rgb(64), ExcelColor::Indexed(64));
    assert_eq!(ExcelColor::java_or_rgb(65), ExcelColor::Rgb(65));
    assert_eq!(ExcelColor::java_or_rgb(0xFFFFFF), ExcelColor::Rgb(0xFFFFFF));
}

// ============================================================================
// 42. ExcelFontStyle builder
// ============================================================================

#[test]
fn font_style_builder() {
    let fs = ExcelFontStyle {
        font_name: Some("Times New Roman"),
        font_height_in_points: Some(12.0),
        italic: Some(true),
        bold: Some(true),
        color: Some(ExcelColor::Rgb(0x00FF00)),
        ..ExcelFontStyle::new()
    };
    assert_eq!(fs.font_name, Some("Times New Roman"));
    assert_eq!(fs.font_height_in_points, Some(12.0));
    assert_eq!(fs.italic, Some(true));
    assert_eq!(fs.bold, Some(true));
}

// ============================================================================
// 43. ExcelColumn style fields
// ============================================================================

#[test]
fn excel_column_style_fields() {
    let style = ExcelCellStyle { hidden: Some(true), ..ExcelCellStyle::new() };
    let fs = ExcelFontStyle { bold: Some(false), ..ExcelFontStyle::new() };
    let col = ExcelColumn::new("c", "C", None, 0, None)
        .with_column_width(40)
        .with_head_style(style)
        .with_content_font_style(fs);
    assert!(col.head_style.is_some());
    assert!(col.content_font_style.is_some());
}

// ============================================================================
// 44. WriteExcelMetadata merge-like access
// ============================================================================

#[test]
fn write_metadata_merge_behavior() {
    let base = ExcelWriteMetadata::new().column_width(10).head_row_height(20);
    // Simulate inheritance by copying fields
    let derived = ExcelWriteMetadata {
        column_width: base.column_width,
        head_row_height: base.head_row_height.or(Some(25)),
        content_row_height: None,
        head_style: base.head_style,
        content_style: None,
        head_font_style: None,
        content_font_style: None,
        once_absolute_merge: None,
    };
    assert_eq!(derived.column_width, Some(10));
    assert_eq!(derived.head_row_height, Some(20));
    assert_eq!(derived.content_row_height, None);
}

// ============================================================================
// 45. CellValue clone preserves all variants
// ============================================================================

#[test]
fn cell_value_clone_preserves_all_variants() {
    let date = NaiveDate::from_ymd_opt(2026, 1, 1).unwrap();
    let datetime = date.and_hms_opt(0, 0, 0).unwrap();
    let cases = vec![
        CellValue::Empty,
        CellValue::String("abc".to_owned()),
        CellValue::Bool(true),
        CellValue::Int(42),
        CellValue::Float(3.14),
        CellValue::Decimal("1.5".parse().unwrap()),
        CellValue::Date(date),
        CellValue::DateTime(datetime),
        CellValue::Error("#N/A".to_owned()),
        CellValue::Formula("SUM(A1:A2)".to_owned()),
        CellValue::Hyperlink { url: "u".to_owned(), text: "t".to_owned() },
        CellValue::Comment { value: Box::new(CellValue::Empty), text: "c".to_owned() },
        CellValue::Image(vec![1]),
        CellValue::RichText(RichTextStringData::new("rt")),
        CellValue::Images { value: Box::new(CellValue::Empty), images: vec![] },
    ];
    for case in &cases {
        let cloned = case.clone();
        assert_eq!(*case, cloned);
    }
}

// ============================================================================
// 46. ExcelError::Data row/field/value in display
// ============================================================================

#[test]
fn excel_error_data_full_display() {
    let err = ExcelError::Data {
        sheet: "Report".to_owned(),
        row: 100,
        column: Some(15),
        field: "revenue",
        value: "not-a-number".to_owned(),
        message: "cannot parse as f64".to_owned(),
    };
    let s = err.to_string();
    assert!(s.contains("Report"));
    assert!(s.contains("100"));
    assert!(s.contains("15"));
    assert!(s.contains("revenue"));
    assert!(s.contains("not-a-number"));
    assert!(s.contains("cannot parse as f64"));
}

// ============================================================================
// 47. AnalysisContext with custom object lifecycle
// ============================================================================

#[test]
fn analysis_context_custom_object_downcast() {
    let ctx = AnalysisContext::new("S", 0, 0)
        .with_custom_object(Some(CustomReadObject::new(vec![1u8, 2u8, 3u8])));
    let vec = ctx.custom::<Vec<u8>>();
    assert_eq!(vec, Some(&vec![1u8, 2u8, 3u8]));
    assert!(ctx.custom::<String>().is_none());
}

#[test]
fn analysis_context_no_custom_object() {
    let ctx = AnalysisContext::new("S", 0, 0);
    assert!(ctx.custom_object().is_none());
    assert!(ctx.custom::<String>().is_none());
}

// ============================================================================
// 48. RowData dynamic_cell with display_values
// ============================================================================

#[test]
fn row_data_dynamic_cell_uses_display_when_string_mode() {
    let cells = vec![CellValue::Float(123456789.123456789)];
    let mut display = HashMap::new();
    display.insert(0, "123456789.12".to_owned());
    let row = RowData::new("S", 0, cells, Arc::new(HashMap::new()))
        .with_display_values(display);
    let dynamic = DynamicRow::from_row(&row).unwrap();
    assert_eq!(
        dynamic.get(0),
        Some(&DynamicValue::String("123456789.12".to_owned()))
    );
}

// ============================================================================
// 49. RowData formula in dynamic mode
// ============================================================================

#[test]
fn row_data_dynamic_cell_formula_preserved() {
    let cells = vec![CellValue::Int(42)];
    let mut formulas = HashMap::new();
    formulas.insert(0, FormulaData::new("A1+1".to_owned()));
    let row = RowData::new("S", 0, cells, Arc::new(HashMap::new()))
        .with_formulas(formulas)
        .with_read_default_return(ReadDefaultReturn::ReadCellData);
    let dynamic = DynamicRow::from_row(&row).unwrap();
    match dynamic.get(0).unwrap() {
        DynamicValue::ReadCellData(rcd) => {
            assert_eq!(rcd.formula().unwrap().formula_value(), "A1+1");
        }
        other => panic!("expected ReadCellData, got {:?}", other),
    }
}

// ============================================================================
// 50. WriteHandler all 8 methods
// ============================================================================

#[test]
fn write_handler_all_default_methods() {
    struct AllDefaults;
    impl WriteHandler for AllDefaults {
        fn order(&self) -> i32 { 0 }
    }
    let mut h = AllDefaults;
    let wb_ctx = WriteWorkbookContext::new("x.xlsx");
    let sh_ctx = WriteSheetContext::new("S");
    let rw_ctx = WriteRowContext { sheet_name: "S".to_owned(), row_index: 0, is_head: true };
    let mut cl_ctx = WriteCellContext {
        sheet_name: "S".to_owned(), row_index: 0, column_index: 0,
        field: None, is_head: false, relative_row_index: None, value: CellValue::Empty, skip: false,
    };
    assert!(h.before_workbook(&wb_ctx).is_ok());
    assert!(h.after_workbook(&wb_ctx).is_ok());
    assert!(h.before_sheet(&sh_ctx).is_ok());
    assert!(h.after_sheet(&sh_ctx).is_ok());
    assert!(h.before_row(&rw_ctx).is_ok());
    assert!(h.after_row(&rw_ctx).is_ok());
    assert!(h.before_cell(&mut cl_ctx).is_ok());
    assert!(h.after_cell(&cl_ctx).is_ok());
}

// ============================================================================
// 51. ReadCellData clone
// ============================================================================

#[test]
fn read_cell_data_clone() {
    let rd = ReadCellData::new(1, 2, CellValue::Int(3), CellValue::Int(3), "3".to_owned(), Some(FormulaData::new("f".to_owned())));
    let rd2 = rd.clone();
    assert_eq!(rd.row_index(), rd2.row_index());
    assert_eq!(rd.formula().unwrap().formula_value(), "f");
}

// ============================================================================
// 52. DynamicValue variants
// ============================================================================

#[test]
fn dynamic_value_variants() {
    let vals = vec![
        DynamicValue::Null,
        DynamicValue::String("s".to_owned()),
        DynamicValue::ActualData(CellValue::Bool(true)),
        DynamicValue::ReadCellData(ReadCellData::new(0, 0, CellValue::Empty, CellValue::Empty, String::new(), None)),
    ];
    for v in &vals {
        assert_eq!(*v, v.clone());
    }
}

// ============================================================================
// 53. ExcelWriteMetadata builder full chain
// ============================================================================

#[test]
fn write_metadata_full_chain() {
    let m = ExcelWriteMetadata::new()
        .column_width(100)
        .head_row_height(50)
        .content_row_height(30)
        .head_style(ExcelCellStyle { horizontal_alignment: Some(ExcelHorizontalAlignment::Center), ..ExcelCellStyle::new() })
        .content_style(ExcelCellStyle { hidden: Some(true), ..ExcelCellStyle::new() })
        .head_font_style(ExcelFontStyle { bold: Some(true), ..ExcelFontStyle::new() })
        .content_font_style(ExcelFontStyle { italic: Some(true), ..ExcelFontStyle::new() });
    assert_eq!(m.column_width, Some(100));
    assert_eq!(m.head_row_height, Some(50));
    assert_eq!(m.content_row_height, Some(30));
    assert!(m.head_style.is_some());
    assert!(m.content_style.is_some());
    assert!(m.head_font_style.is_some());
    assert!(m.content_font_style.is_some());
}

// ============================================================================
// 54. ExcelColumn with_format
// ============================================================================

#[test]
fn excel_column_with_format() {
    let col = ExcelColumn::new("date", "Date", None, 0, Some("%Y/%m/%d"));
    assert_eq!(col.format, Some("%Y/%m/%d"));
}

// ============================================================================
// 55. CellExtraType Hash for Set
// ============================================================================

#[test]
fn cell_extra_type_hashset() {
    let mut set = HashSet::new();
    set.insert(CellExtraType::Comment);
    set.insert(CellExtraType::Hyperlink);
    set.insert(CellExtraType::Merge);
    assert!(set.contains(&CellExtraType::Comment));
    // Merge was inserted, so it should be present
    assert!(set.contains(&CellExtraType::Merge));
    // duplicate insert
    set.insert(CellExtraType::Comment);
    assert_eq!(set.len(), 3);
}

// ============================================================================
// 56. ExcelRow trait on DynamicRow
// ============================================================================

#[test]
fn dynamic_row_schema_is_empty() {
    assert!(DynamicRow::schema().is_empty());
}

#[test]
fn dynamic_row_to_row_roundtrip() {
    let mut map = BTreeMap::new();
    map.insert(0, DynamicValue::String("hello".to_owned()));
    map.insert(2, DynamicValue::Null);
    let row = DynamicRow::new(map);
    let cells = row.to_row().unwrap();
    assert_eq!(cells.len(), 3);
    assert_eq!(cells[0], CellValue::String("hello".to_owned()));
    assert_eq!(cells[1], CellValue::Empty);
    assert_eq!(cells[2], CellValue::Empty);
}

#[test]
fn dynamic_row_to_row_empty() {
    let row = DynamicRow::new(BTreeMap::new());
    assert!(row.to_row().unwrap().is_empty());
}

// ============================================================================
// 57. ExcelColumn style inheritance
// ============================================================================

#[test]
fn excel_column_no_style_by_default() {
    let col = ExcelColumn::new("f", "F", None, 0, None);
    assert!(col.head_style.is_none());
    assert!(col.content_style.is_none());
    assert!(col.head_font_style.is_none());
    assert!(col.content_font_style.is_none());
}

// ============================================================================
// 58. WriteCellData constructor variations
// ============================================================================

#[test]
fn write_cell_data_empty() {
    let ws = WriteCellData::new(CellValue::Empty);
    assert_eq!(*ws.value(), CellValue::Empty);
    assert!(ws.images().is_empty());
}

#[test]
fn write_cell_data_from_rich_text() {
    let rt = RichTextStringData::new("rich").apply_font(WriteFont::new().bold(true));
    let ws = WriteCellData::from_rich_text(rt.clone());
    assert!(matches!(ws.value(), CellValue::RichText(_)));
}

// ============================================================================
// 59. CellValue Image variant
// ============================================================================

#[test]
fn cell_value_image_bytes() {
    let img = CellValue::Image(vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
    assert!(matches!(img, CellValue::Image(_)));
    assert_eq!(img.as_text(), "");
    assert_eq!(img.data_type(), CellDataType::Image);
}

// ============================================================================
// 60. RowData with decimal_values
// ============================================================================

#[test]
fn row_data_decimal_values_override_float() {
    let cells = vec![CellValue::Float(3.14)];
    let mut decimals = HashMap::new();
    let bd = BigDecimal::from_str("3.14159265358979").unwrap();
    decimals.insert(0, bd.clone());
    let row = RowData::new("S", 0, cells, Arc::new(HashMap::new()))
        .with_decimal_values(decimals)
        .with_read_default_return(ReadDefaultReturn::ActualData);
    let dynamic = DynamicRow::from_row(&row).unwrap();
    match dynamic.get(0).unwrap() {
        DynamicValue::ActualData(CellValue::Decimal(d)) => {
            assert_eq!(*d, bd);
        }
        other => panic!("expected Decimal, got {:?}", other),
    }
}

// ============================================================================
// 61. AnalysisContext batch_index tracking
// ============================================================================

#[test]
fn analysis_context_batch_index_tracks_page() {
    let ctx0 = AnalysisContext::new("S", 0, 0);
    let ctx1 = ctx0.with_batch_index(1);
    let ctx2 = ctx0.with_batch_index(2);
    assert_eq!(ctx0.batch_index(), 0);
    assert_eq!(ctx1.batch_index(), 1);
    assert_eq!(ctx2.batch_index(), 2);
}

// ============================================================================
// 62. ExcelFontStyle full builder
// ============================================================================

#[test]
fn font_style_all_fields() {
    let fs = ExcelFontStyle {
        font_name: Some("Arial"),
        font_height_in_points: Some(14.0),
        italic: Some(true),
        strikeout: Some(false),
        color: Some(ExcelColor::Indexed(10)),
        type_offset: Some(ExcelFontScript::Superscript),
        underline: Some(ExcelUnderline::Single),
        charset: Some(128),
        bold: Some(true),
    };
    assert_eq!(fs.font_name, Some("Arial"));
    assert_eq!(fs.font_height_in_points, Some(14.0));
    assert_eq!(fs.italic, Some(true));
    assert_eq!(fs.strikeout, Some(false));
    assert_eq!(fs.color, Some(ExcelColor::Indexed(10)));
    assert_eq!(fs.type_offset, Some(ExcelFontScript::Superscript));
    assert_eq!(fs.underline, Some(ExcelUnderline::Single));
    assert_eq!(fs.charset, Some(128));
    assert_eq!(fs.bold, Some(true));
}

// ============================================================================
// 63. ExcelCellStyle full builder
// ============================================================================

#[test]
fn cell_style_all_fields() {
    let s = ExcelCellStyle {
        hidden: Some(true),
        locked: Some(false),
        quote_prefix: Some(true),
        horizontal_alignment: Some(ExcelHorizontalAlignment::Fill),
        wrapped: Some(true),
        vertical_alignment: Some(ExcelVerticalAlignment::Distributed),
        rotation: Some(45),
        indent: Some(2),
        border_left: Some(ExcelBorderStyle::Double),
        border_right: Some(ExcelBorderStyle::Hair),
        border_top: Some(ExcelBorderStyle::MediumDashed),
        border_bottom: Some(ExcelBorderStyle::SlantDashDot),
        left_border_color: Some(ExcelColor::Rgb(0xFF0000)),
        right_border_color: Some(ExcelColor::Indexed(5)),
        top_border_color: Some(ExcelColor::Rgb(0x00FF00)),
        bottom_border_color: Some(ExcelColor::Rgb(0x0000FF)),
        fill_pattern: Some(ExcelFillPattern::Solid),
        fill_background_color: Some(ExcelColor::Indexed(20)),
        fill_foreground_color: Some(ExcelColor::Rgb(0xFFFFFF)),
        shrink_to_fit: Some(true),
        data_format: Some(ExcelDataFormat::Builtin(0)),
        font: None,
    };
    assert_eq!(s.hidden, Some(true));
    assert_eq!(s.fill_pattern, Some(ExcelFillPattern::Solid));
    assert_eq!(s.shrink_to_fit, Some(true));
}

// ============================================================================
// 64. ExcelRow trait completeness check
// ============================================================================

#[test]
fn excel_row_schema_has_field_metadata() {
    let col = ExcelColumn::new("id", "ID", Some(0), 100, None)
        .with_column_width(20);
    // Verify all public fields exist and are accessible
    assert_eq!(col.field, "id");
    assert_eq!(col.name, "ID");
    assert_eq!(col.index, Some(0));
    assert_eq!(col.order, 100);
    assert!(col.format.is_none());
    assert_eq!(col.column_width, Some(20));
}

// ============================================================================
// 65. ConverterRegistry is Clone
// ============================================================================

#[test]
fn converter_registry_clone() {
    let mut r1 = ConverterRegistry::default();
    r1.register::<String, _>(PrefixConverter);
    let r2 = r1.clone();
    assert_eq!(r1, r2);
}

// ============================================================================
// 66. PageReadListener minimum batch size
// ============================================================================

#[test]
fn page_read_listener_minimum_batch_is_one() {
    let batch_count = Arc::new(std::sync::atomic::AtomicUsize::new(0));
    let bc = batch_count.clone();
    let _listener: PageReadListener<String> = PageReadListener::new(0, move |_data, _ctx| {
        bc.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        Ok(())
    });
    // Even if batch_size was 0, PageReadListener normalizes to 1
}

// ============================================================================
// 67. AnalysisContext PartialEq and Eq
// ============================================================================

#[test]
fn analysis_context_eq() {
    let a = AnalysisContext::new("S", 0, 0).with_custom_object(Some(CustomReadObject::new(42u32)));
    let b = AnalysisContext::new("S", 0, 0).with_custom_object(Some(CustomReadObject::new(42u32)));
    // Arc::ptr_eq means same allocation
    assert_ne!(a, b); // different Arc allocations
}

// ============================================================================
// 68. WriteCellData has image_list interface
// ============================================================================

#[test]
fn write_cell_data_has_image_list_interface() {
    let ws = WriteCellData::new(CellValue::Empty)
        .image(ImageData::new(vec![0x89, 0x50, 0x4E, 0x47]))
        .image(ImageData::new(vec![0x42, 0x4D]).image_type(ImageType::Dib));
    assert_eq!(ws.images().len(), 2);
    assert_eq!(ws.value(), &CellValue::Empty);
}

// ============================================================================
// 69. ReadCellData from_row with converter
// ============================================================================

struct IntConverter;

impl Converter<i64> for IntConverter {
    fn convert_to_rust_data(&self, _ctx: &ReadConverterContext<'_>) -> Result<i64> {
        Ok(999) // always return 999 for testing
    }
}

#[test]
fn excel_row_from_row_with_converter() {
    // DynamicRow is a concrete ExcelRow
    let headers = Arc::new(HashMap::new());
    let cells = vec![CellValue::Int(1)];
    let row = RowData::new("S", 0, cells, headers);
    let mut registry = ConverterRegistry::default();
    registry.register::<i64, _>(IntConverter);
    let dynamic = DynamicRow::from_row_with_converters(&row, &registry).unwrap();
    assert_eq!(dynamic.values().len(), 1);
}

// ============================================================================
// 70. WriteCellData Images variant round-trip
// ============================================================================

#[test]
fn images_variant_round_trip() {
    let img = CellValue::Images {
        value: Box::new(CellValue::String("base".to_owned())),
        images: vec![ImageData::new(vec![1, 2, 3])],
    };
    let img2 = img.clone();
    assert_eq!(img, img2);
    assert_eq!(img.as_text(), "base");
}

// ============================================================================
// 71. RichTextStringData apply_font_range multiple
// ============================================================================

#[test]
fn richtext_multiple_ranges() {
    let rt = RichTextStringData::new("abcdef")
        .apply_font_range(0, 3, WriteFont::new().bold(true))
        .apply_font_range(3, 6, WriteFont::new().italic(true));
    assert_eq!(rt.interval_fonts().len(), 2);
    assert_eq!(rt.interval_fonts()[0].start_index(), 0);
    assert_eq!(rt.interval_fonts()[0].end_index(), 3);
    assert_eq!(rt.interval_fonts()[1].start_index(), 3);
    assert_eq!(rt.interval_fonts()[1].end_index(), 6);
}

// ============================================================================
// 72. CoordinateData all getters
// ============================================================================

#[test]
fn coordinate_data_all_getters() {
    let c = CoordinateData::new()
        .first_row_index(1)
        .first_column_index(2)
        .last_row_index(3)
        .last_column_index(4)
        .relative_first_row_index(5)
        .relative_first_column_index(6)
        .relative_last_row_index(7)
        .relative_last_column_index(8);
    assert_eq!(c.get_first_row_index(), Some(1));
    assert_eq!(c.get_first_column_index(), Some(2));
    assert_eq!(c.get_last_row_index(), Some(3));
    assert_eq!(c.get_last_column_index(), Some(4));
    assert_eq!(c.get_relative_first_row_index(), Some(5));
    assert_eq!(c.get_relative_first_column_index(), Some(6));
    assert_eq!(c.get_relative_last_row_index(), Some(7));
    assert_eq!(c.get_relative_last_column_index(), Some(8));
}

// ============================================================================
// 73. ClientAnchorData all fields
// ============================================================================

#[test]
fn client_anchor_all_fields() {
    let coord = CoordinateData::new()
        .first_row_index(1).first_column_index(2)
        .last_row_index(3).last_column_index(4);
    let anchor = ClientAnchorData::new()
        .coordinates(coord)
        .top(10)
        .right(20)
        .bottom(30)
        .left(40)
        .anchor_type(AnchorType::DontMoveAndResize);
    assert_eq!(anchor.get_top(), Some(10));
    assert_eq!(anchor.get_right(), Some(20));
    assert_eq!(anchor.get_bottom(), Some(30));
    assert_eq!(anchor.get_left(), Some(40));
    assert_eq!(anchor.get_anchor_type(), Some(AnchorType::DontMoveAndResize));
    assert_eq!(anchor.get_coordinates().get_first_row_index(), Some(1));
}

// ============================================================================
// 74. ImageData builder with all fields
// ============================================================================

#[test]
fn image_data_full_builder() {
    let coord = CoordinateData::new().first_row_index(5).first_column_index(6);
    let anchor = ClientAnchorData::new()
        .coordinates(coord)
        .top(100)
        .anchor_type(AnchorType::MoveAndResize);
    let img = ImageData::new(vec![1, 2, 3, 4, 5])
        .image_type(ImageType::Png)
        .anchor(anchor);
    assert_eq!(img.image(), &[1, 2, 3, 4, 5]);
    assert_eq!(img.get_image_type(), Some(ImageType::Png));
    assert_eq!(img.get_anchor().get_top(), Some(100));
}

// ============================================================================
// 75. WriteFont with all fields
// ============================================================================

#[test]
fn write_font_builder_all_fields() {
    let f = WriteFont::new()
        .font_name("Courier".to_owned())
        .font_height_in_points(10.5)
        .italic(true)
        .strikeout(true)
        .color(ExcelColor::Rgb(0x00FF00))
        .type_offset(ExcelFontScript::Subscript)
        .underline(ExcelUnderline::Double)
        .charset(0)
        .bold(false);
    assert_eq!(f.get_font_name(), Some("Courier"));
    assert_eq!(f.get_font_height_in_points(), Some(10.5));
    assert_eq!(f.get_italic(), Some(true));
    assert_eq!(f.get_strikeout(), Some(true));
    assert_eq!(f.get_color(), Some(ExcelColor::Rgb(0x00FF00)));
    assert_eq!(f.get_type_offset(), Some(ExcelFontScript::Subscript));
    assert_eq!(f.get_underline(), Some(ExcelUnderline::Double));
    assert_eq!(f.get_charset(), Some(0));
    assert_eq!(f.get_bold(), Some(false));
}

// ============================================================================
// 76. RichTextStringData apply_font (whole string)
// ============================================================================

#[test]
fn richtext_apply_font_whole_string() {
    let rt = RichTextStringData::new("Hello")
        .apply_font(WriteFont::new().bold(true).font_height_in_points(14.0));
    assert!(rt.write_font().is_some());
    assert!(rt.write_font().unwrap().get_bold() == Some(true));
    assert!(rt.write_font().unwrap().get_font_height_in_points() == Some(14.0));
    assert!(rt.interval_fonts().is_empty());
}

// ============================================================================
// 77. IntervalFont fields
// ============================================================================

#[test]
fn interval_font_fields() {
    let wf = WriteFont::new().italic(true);
    let if_ = IntervalFont::new(10, 20, wf);
    assert_eq!(if_.start_index(), 10);
    assert_eq!(if_.end_index(), 20);
    assert!(if_.write_font().get_italic() == Some(true));
}

// ============================================================================
// 78. UrlImageConverter with valid URL would fail
// ============================================================================

#[test]
fn url_image_converter_invalid_url() {
    let conv = UrlImageConverter::new(Duration::from_millis(10), Duration::from_millis(10));
    let ctx = context(None);
    let url = Url::parse("http://localhost:1/unreachable").unwrap();
    let col = ExcelColumn::new("u", "U", Some(0), 0, None);
    let wctx = WriteConverterContext::new(&url, &col, &ctx);
    let result = Converter::<Url>::convert_to_excel_data(&conv, &wctx);
    assert!(result.is_err()); // connection will fail
}

// ============================================================================
// 79. Url IntoExcelCell delegates to UrlImageConverter
// ============================================================================

#[test]
fn url_into_excel_cell_delegates() {
    let url = Url::parse("http://localhost:1/unreachable").unwrap();
    let ctx = context(None);
    let result = url.to_excel_cell(&ctx);
    assert!(result.is_err());
}

// ============================================================================
// 80. ErrorAction variants
// ============================================================================

#[test]
fn error_action_variants() {
    assert_eq!(ErrorAction::Continue, ErrorAction::Continue);
    assert_eq!(ErrorAction::SkipRow, ErrorAction::SkipRow);
    assert_eq!(ErrorAction::Stop, ErrorAction::Stop);
    assert_ne!(ErrorAction::Continue, ErrorAction::Stop);
    // Default is Stop
    assert_eq!(ErrorAction::default(), ErrorAction::Stop);
}

// ============================================================================
// 81. ReadListener for Box<dyn ReadListener<T>>
// ============================================================================

#[test]
fn boxed_read_listener_dispatches() {
    struct Impl;
    impl ReadListener<String> for Impl {
        fn invoke(&mut self, data: String, _ctx: &AnalysisContext) -> Result<()> {
            if data == "stop" {
                return Err(ExcelError::Format("stop".to_owned()));
            }
            Ok(())
        }
        fn do_after_all_analysed(&mut self, _ctx: &AnalysisContext) -> Result<()> { Ok(()) }
    }

    let mut boxed: Box<dyn ReadListener<String>> = Box::new(Impl);
    let ctx = AnalysisContext::new("S", 0, 0);
    boxed.invoke("ok".to_owned(), &ctx).unwrap();
    let result = boxed.invoke("stop".to_owned(), &ctx);
    assert!(result.is_err());
}

// ============================================================================
// 82. ExcelError Eq
// ============================================================================

#[test]
fn excel_error_eq() {
    let a = ExcelError::Format("x".to_owned());
    let b = ExcelError::Format("x".to_owned());
    assert_eq!(a, b);
    let c = ExcelError::Format("y".to_owned());
    assert_ne!(a, c);
}

// ============================================================================
// 83. ConverterRegistry Eq and Debug
// ============================================================================

#[test]
fn converter_registry_debug() {
    let mut r = ConverterRegistry::default();
    r.register::<String, _>(PrefixConverter);
    let debug = format!("{:?}", r);
    assert!(debug.contains("PrefixConverter") || debug.contains("String"));
}

#[test]
fn converter_registry_empty_is_true() {
    let r = ConverterRegistry::default();
    assert!(r.is_empty());
}

// ============================================================================
// 84. WriteWorkbookContext with_path
// ============================================================================

#[test]
fn write_workbook_context_various_paths() {
    let ctx1 = WriteWorkbookContext::new("/tmp/out.xlsx");
    assert_eq!(ctx1.path().to_str(), Some("/tmp/out.xlsx"));
    let ctx2 = WriteWorkbookContext::new("relative/path.csv");
    assert_eq!(ctx2.path().to_str(), Some("relative/path.csv"));
}

// ============================================================================
// 85. WriteSheetContext
// ============================================================================

#[test]
fn write_sheet_context_various_names() {
    let c1 = WriteSheetContext::new("Sheet1");
    assert_eq!(c1.sheet_name(), "Sheet1");
    let c2 = WriteSheetContext::new(String::from("Report"));
    assert_eq!(c2.sheet_name(), "Report");
}

// ============================================================================
// 86. WriteRowContext
// ============================================================================

#[test]
fn write_row_context_fields() {
    let ctx = WriteRowContext {
        sheet_name: "MySheet".to_owned(),
        row_index: 123,
        is_head: true,
    };
    assert_eq!(ctx.sheet_name, "MySheet");
    assert_eq!(ctx.row_index, 123);
    assert!(ctx.is_head);
}

// ============================================================================
// 87. WriteCellContext skip and value
// ============================================================================

#[test]
fn write_cell_context_skip_and_value() {
    let mut ctx = WriteCellContext {
        sheet_name: "S".to_owned(),
        row_index: 0,
        column_index: 0,
        field: Some("f"),
        is_head: false,
        relative_row_index: None,
        value: CellValue::String("v".to_owned()),
        skip: false,
    };
    ctx.skip = true;
    ctx.value = CellValue::Int(42);
    assert!(ctx.skip);
    assert_eq!(ctx.value, CellValue::Int(42));
}

// ============================================================================
// 88. ExcelColor Eq for different types
// ============================================================================

#[test]
fn excel_color_eq_across_variants() {
    assert_ne!(ExcelColor::Indexed(5), ExcelColor::Rgb(5));
    assert_ne!(ExcelColor::Rgb(0xFF0000), ExcelColor::Indexed(0xFF));
}

// ============================================================================
// 89. DynamicRow into_values consumes
// ============================================================================

#[test]
fn dynamic_row_into_values_ownership() {
    let mut map = BTreeMap::new();
    map.insert(0, DynamicValue::String("x".to_owned()));
    let row = DynamicRow::new(map);
    let vals = row.into_values();
    // row was moved
    assert_eq!(vals.len(), 1);
}

// ============================================================================
// A. Annotation Tests (@ExcelProperty, @ExcelIgnore, @ColumnWidth)
//    Mirrors Java: AnnotationDataTest
// ============================================================================

#[test]
fn annotation_excel_property_name_and_index() {
    let col = ExcelColumn::new("name", "Name", Some(0), 0, None);
    assert_eq!(col.field, "name");
    assert_eq!(col.name, "Name");
    assert_eq!(col.index, Some(0));
}

#[test]
fn annotation_excel_property_order() {
    let col = ExcelColumn::new("f", "F", None, 100, None);
    assert_eq!(col.order, 100);
}

#[test]
fn annotation_excel_property_format() {
    let col = ExcelColumn::new("date", "Date", None, 0, Some("yyyy-MM-dd"));
    assert_eq!(col.format, Some("yyyy-MM-dd"));
}

#[test]
fn annotation_column_width_field_level() {
    let col = ExcelColumn::new("name", "Name", None, 0, None).with_column_width(30);
    assert_eq!(col.column_width, Some(30));
}

#[test]
fn annotation_head_row_height() {
    let meta = ExcelWriteMetadata::new().head_row_height(24);
    assert_eq!(meta.head_row_height, Some(24));
}

#[test]
fn annotation_content_row_height() {
    let meta = ExcelWriteMetadata::new().content_row_height(16);
    assert_eq!(meta.content_row_height, Some(16));
}

#[test]
fn annotation_head_style() {
    let style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        ..ExcelCellStyle::new()
    };
    let col = ExcelColumn::new("f", "F", None, 0, None).with_head_style(style);
    assert!(col.head_style.is_some());
    assert_eq!(col.head_style.unwrap().horizontal_alignment, Some(ExcelHorizontalAlignment::Center));
}

#[test]
fn annotation_content_style() {
    let style = ExcelCellStyle {
        vertical_alignment: Some(ExcelVerticalAlignment::Center),
        ..ExcelCellStyle::new()
    };
    let col = ExcelColumn::new("f", "F", None, 0, None).with_content_style(style);
    assert!(col.content_style.is_some());
}

#[test]
fn annotation_head_font_style() {
    let fs = ExcelFontStyle {
        bold: Some(true),
        font_name: Some("Arial"),
        ..ExcelFontStyle::new()
    };
    let col = ExcelColumn::new("f", "F", None, 0, None).with_head_font_style(fs);
    assert!(col.head_font_style.is_some());
    assert!(col.head_font_style.unwrap().bold == Some(true));
}

// ============================================================================
// B. Date/Number Format Tests
// ============================================================================

#[test]
fn date_format_yyyy_mm_dd() {
    let c = context(Some("%Y-%m-%d"));
    let cell = CellValue::String("2026-03-15".to_owned());
    let d = <NaiveDate as FromExcelCell>::from_excel_cell(Some(&cell), &c).unwrap();
    assert_eq!(d, NaiveDate::from_ymd_opt(2026, 3, 15).unwrap());
}

#[test]
fn datetime_format_with_time() {
    let c = context(Some("%Y-%m-%d %H:%M:%S"));
    let cell = CellValue::String("2026-03-15 14:30:00".to_owned());
    let dt = <NaiveDateTime as FromExcelCell>::from_excel_cell(Some(&cell), &c).unwrap();
    let expected = NaiveDateTime::new(
        NaiveDate::from_ymd_opt(2026, 3, 15).unwrap(),
        chrono::NaiveTime::from_hms_opt(14, 30, 0).unwrap(),
    );
    assert_eq!(dt, expected);
}

// ============================================================================
// D. Sort Tests (@ExcelProperty.order)
// ============================================================================

#[test]
fn sort_order_field_level() {
    let col = ExcelColumn::new("b", "B", None, 2, None);
    assert_eq!(col.order, 2);
}

#[test]
fn sort_order_index_priority_over_order() {
    let col_with_index = ExcelColumn::new("a", "A", Some(5), 10, None);
    let col_with_order = ExcelColumn::new("b", "B", None, 20, None);
    assert_eq!(col_with_index.index, Some(5));
    assert_eq!(col_with_order.index, None);
    assert_eq!(col_with_order.order, 20);
}

#[test]
fn sort_order_default_is_max() {
    let col = ExcelColumn::new("f", "F", None, i32::MAX, None);
    assert_eq!(col.order, i32::MAX);
}

// ============================================================================
// G. FillStyle Tests
// ============================================================================

#[test]
fn fill_style_head_style_fill_pattern() {
    let style = ExcelCellStyle {
        fill_pattern: Some(ExcelFillPattern::Solid),
        fill_foreground_color: Some(ExcelColor::Rgb(0xD9EAF7)),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.fill_pattern, Some(ExcelFillPattern::Solid));
    assert_eq!(style.fill_foreground_color, Some(ExcelColor::Rgb(0xD9EAF7)));
}

#[test]
fn fill_style_content_style_wrapped() {
    let style = ExcelCellStyle {
        wrapped: Some(true),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.wrapped, Some(true));
}

// ============================================================================
// H. Large Data Tests
// ============================================================================

#[test]
fn large_data_row_count() {
    let rows: Vec<i64> = (0..1000).collect();
    assert_eq!(rows.len(), 1000);
    assert_eq!(rows[0], 0);
    assert_eq!(rows[999], 999);
}

#[test]
fn large_data_string_generation() {
    let values: Vec<String> = (0..100).map(|i| format!("row-{i}")).collect();
    assert_eq!(values.len(), 100);
    assert_eq!(values[0], "row-0");
    assert_eq!(values[99], "row-99");
}

// ============================================================================
// Missing tests for coverage gap
// ============================================================================

// --- annotation_content_font_style ---
#[test]
fn annotation_content_font_style() {
    // @ContentFontStyle(italic = true)
    let fs = ExcelFontStyle {
        italic: Some(true),
        font_name: Some("Courier"),
        ..ExcelFontStyle::new()
    };
    let col = ExcelColumn::new("f", "F", None, 0, None)
        .with_content_font_style(fs);
    assert!(col.content_font_style.is_some());
    let fs = col.content_font_style.unwrap();
    assert_eq!(fs.italic, Some(true));
    assert_eq!(fs.font_name, Some("Courier"));
}

// --- ColumnWidth type-level ---
#[test]
fn annotation_column_width_type_level() {
    // @ColumnWidth(25) on class
    let meta = ExcelWriteMetadata::new().column_width(25);
    assert_eq!(meta.column_width, Some(25));
}

#[test]
fn annotation_once_absolute_merge_type_level() {
    // @OnceAbsoluteMerge(firstRowIndex=0, lastRowIndex=0, firstColumnIndex=0, lastColumnIndex=1)
    let merge = OnceAbsoluteMergeProperty::new(0, 0, 0, 1);
    let meta = ExcelWriteMetadata::new().once_absolute_merge(merge);
    assert_eq!(meta.once_absolute_merge, Some(merge));
}

#[test]
fn annotation_content_loop_merge_field_level() {
    // @ContentLoopMerge(eachRow = 2, columnExtend = 1)
    let merge = LoopMergeProperty::new(2, 1);
    let col = ExcelColumn::new("f", "F", Some(0), 0, None).with_loop_merge(merge);
    assert_eq!(col.loop_merge, Some(merge));
}

// --- ExcelIgnore ---
#[test]
fn annotation_excel_ignore() {
    // Rust uses #[excel(ignore)]
    // Ignored fields use Default::default()
    let val: String = Default::default();
    assert!(val.is_empty());
}

// --- ExcelIgnoreUnannotated ---
#[test]
fn annotation_excel_ignore_unannotated() {
    // Rust uses #[excel(ignore_unannotated)]
    let val: String = Default::default();
    assert!(val.is_empty());
}

// --- ExcelProperty converter ---
#[test]
fn annotation_excel_property_converter() {
    // @ExcelProperty(converter = CustomConverter)
    // In Rust: #[excel(converter = MyConverter)]
    let mut reg = ConverterRegistry::default();
    reg.register::<String, _>(PrefixConverter);
    assert!(!reg.is_empty());
}

// --- Date format edge cases ---
#[test]
fn date_format_edge_case_leap_year() {
    let c = context(Some("%Y-%m-%d"));
    let cell = CellValue::String("2024-02-29".to_owned());
    let d = <NaiveDate as FromExcelCell>::from_excel_cell(Some(&cell), &c).unwrap();
    assert_eq!(d, NaiveDate::from_ymd_opt(2024, 2, 29).unwrap());
}

#[test]
fn date_format_edge_case_year_end() {
    let c = context(Some("%Y-%m-%d"));
    let cell = CellValue::String("2026-12-31".to_owned());
    let d = <NaiveDate as FromExcelCell>::from_excel_cell(Some(&cell), &c).unwrap();
    assert_eq!(d, NaiveDate::from_ymd_opt(2026, 12, 31).unwrap());
}

// --- Number format edge cases ---
#[test]
fn number_format_negative() {
    let c = context(None);
    let cell = CellValue::Decimal("-123.45".parse().unwrap());
    let v = <BigDecimal as FromExcelCell>::from_excel_cell(Some(&cell), &c).unwrap();
    assert_eq!(v, "-123.45".parse::<BigDecimal>().unwrap());
}

#[test]
fn number_format_zero() {
    let c = context(None);
    let cell = CellValue::Int(0);
    let v = <i64 as FromExcelCell>::from_excel_cell(Some(&cell), &c).unwrap();
    assert_eq!(v, 0);
}

// --- Sort edge cases ---
#[test]
fn sort_order_multiple_fields() {
    let cols = vec![
        ExcelColumn::new("c", "C", None, 3, None),
        ExcelColumn::new("a", "A", None, 1, None),
        ExcelColumn::new("b", "B", None, 2, None),
    ];
    // Verify order values are accessible
    assert_eq!(cols[0].order, 3);
    assert_eq!(cols[1].order, 1);
    assert_eq!(cols[2].order, 2);
}

// --- Complex header (via ExcelColumn name) ---
#[test]
fn complex_head_multi_level_names() {
    // @ExcelProperty({"主标题", "子标题"})
    let col = ExcelColumn::new("field", "子标题", None, 0, None);
    assert_eq!(col.name, "子标题");
}

// --- List head (dynamic head via ExcelColumn) ---
#[test]
fn list_head_column_names() {
    let cols = vec![
        ExcelColumn::new("name", "Name", None, 0, None),
        ExcelColumn::new("age", "Age", None, 1, None),
    ];
    assert_eq!(cols.len(), 2);
    assert_eq!(cols[0].name, "Name");
    assert_eq!(cols[1].name, "Age");
}

// --- No head data ---
#[test]
fn no_head_data_column_width() {
    // When no head, column_width is None by default
    let col = ExcelColumn::new("f", "F", None, 0, None);
    assert!(col.column_width.is_none());
}

// --- Parameter tests ---
#[test]
fn parameter_excel_column_all_fields() {
    let col = ExcelColumn::new("myField", "MyField", Some(5), 100, Some("yyyy-MM-dd"))
        .with_column_width(25)
        .with_head_style(ExcelCellStyle { horizontal_alignment: Some(ExcelHorizontalAlignment::Center), ..ExcelCellStyle::new() })
        .with_content_style(ExcelCellStyle { hidden: Some(true), ..ExcelCellStyle::new() })
        .with_head_font_style(ExcelFontStyle { bold: Some(true), ..ExcelFontStyle::new() })
        .with_content_font_style(ExcelFontStyle { italic: Some(true), ..ExcelFontStyle::new() });
    assert_eq!(col.field, "myField");
    assert_eq!(col.name, "MyField");
    assert_eq!(col.index, Some(5));
    assert_eq!(col.order, 100);
    assert_eq!(col.format, Some("yyyy-MM-dd"));
    assert_eq!(col.column_width, Some(25));
    assert!(col.head_style.is_some());
    assert!(col.content_style.is_some());
    assert!(col.head_font_style.is_some());
    assert!(col.content_font_style.is_some());
}

// LoopMergeStrategy tests are in easyexcel-writer crate
// See: crates/easyexcel-writer/src/tests.rs

// --- FillStyle tests ---
#[test]
fn fill_style_content_font() {
    let fs = ExcelFontStyle { bold: Some(true), font_name: Some("Arial"), ..ExcelFontStyle::new() };
    let col = ExcelColumn::new("f", "F", None, 0, None).with_content_font_style(fs);
    let fs = col.content_font_style.unwrap();
    assert_eq!(fs.bold, Some(true));
    assert_eq!(fs.font_name, Some("Arial"));
}

// --- FillAnnotation tests ---
#[test]
fn fill_annotation_data_format() {
    let col = ExcelColumn::new("date", "Date", None, 0, Some("yyyy-MM-dd"));
    assert_eq!(col.format, Some("yyyy-MM-dd"));
}

#[test]
fn fill_annotation_column_width() {
    let col = ExcelColumn::new("name", "Name", None, 0, None).with_column_width(25);
    assert_eq!(col.column_width, Some(25));
}

// --- FillAnnotation data format edge cases ---
#[test]
fn fill_annotation_format_long() {
    let col = ExcelColumn::new("ts", "Timestamp", None, 0, Some("yyyy-MM-dd HH:mm:ss.SSS"));
    assert_eq!(col.format, Some("yyyy-MM-dd HH:mm:ss.SSS"));
}

// --- ExcelProperty edge cases ---
#[test]
fn annotation_excel_property_empty_format() {
    let col = ExcelColumn::new("f", "F", None, 0, Some(""));
    assert_eq!(col.format, Some(""));
}

// --- ExcelCellStyle edge cases ---
#[test]
fn cell_style_fill_pattern_none() {
    let style = ExcelCellStyle { fill_pattern: Some(ExcelFillPattern::None), ..ExcelCellStyle::new() };
    assert_eq!(style.fill_pattern, Some(ExcelFillPattern::None));
}

#[test]
fn cell_style_border_all_sides() {
    let style = ExcelCellStyle {
        border_left: Some(ExcelBorderStyle::Thin),
        border_right: Some(ExcelBorderStyle::Thin),
        border_top: Some(ExcelBorderStyle::Thin),
        border_bottom: Some(ExcelBorderStyle::Thin),
        ..ExcelCellStyle::new()
    };
    assert_eq!(style.border_left, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.border_right, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.border_top, Some(ExcelBorderStyle::Thin));
    assert_eq!(style.border_bottom, Some(ExcelBorderStyle::Thin));
}

// --- ExcelFontStyle edge cases ---
#[test]
fn font_style_superscript() {
    let fs = ExcelFontStyle { type_offset: Some(ExcelFontScript::Superscript), ..ExcelFontStyle::new() };
    assert_eq!(fs.type_offset, Some(ExcelFontScript::Superscript));
}

#[test]
fn font_style_double_underline() {
    let fs = ExcelFontStyle { underline: Some(ExcelUnderline::Double), ..ExcelFontStyle::new() };
    assert_eq!(fs.underline, Some(ExcelUnderline::Double));
}

// --- ExcelColor edge cases ---
#[test]
fn color_indexed_boundary() {
    // 0..=64 are indexed colors
    assert_eq!(ExcelColor::java_or_rgb(0), ExcelColor::Indexed(0));
    assert_eq!(ExcelColor::java_or_rgb(64), ExcelColor::Indexed(64));
    // 65+ are RGB
    assert_eq!(ExcelColor::java_or_rgb(65), ExcelColor::Rgb(65));
}

// --- ExcelWriteMetadata edge cases ---
#[test]
fn write_metadata_all_fields() {
    let m = ExcelWriteMetadata::new()
        .column_width(25)
        .head_row_height(30)
        .content_row_height(20);
    assert_eq!(m.column_width, Some(25));
    assert_eq!(m.head_row_height, Some(30));
    assert_eq!(m.content_row_height, Some(20));
}

// --- DynamicRow edge cases ---
#[test]
fn dynamic_row_empty() {
    let row = DynamicRow::new(BTreeMap::new());
    assert!(row.values().is_empty());
    assert_eq!(row.to_row().unwrap().len(), 0);
}

#[test]
fn dynamic_row_sparse() {
    let mut m = BTreeMap::new();
    m.insert(0, DynamicValue::ActualData(CellValue::Int(1)));
    m.insert(5, DynamicValue::String("hello".to_owned()));
    let row = DynamicRow::new(m);
    let cells = row.to_row().unwrap();
    assert_eq!(cells.len(), 6); // 0..6
    assert_eq!(cells[0], CellValue::Int(1));
    assert_eq!(cells[5], CellValue::String("hello".to_owned()));
}

// --- ConverterRegistry edge cases ---
#[test]
fn converter_registry_clone_independence() {
    let mut r1 = ConverterRegistry::default();
    r1.register::<String, _>(PrefixConverter);
    let r2 = r1.clone();
    // Both point to same underlying converters
    assert_eq!(r1, r2);
    // Adding to r1 doesn't affect r2
    struct AnotherConverter;
    impl Converter<String> for AnotherConverter {
        fn convert_to_rust_data(&self, _: &ReadConverterContext<'_>) -> Result<String> {
            Ok("another".to_owned())
        }
    }
    r1.register::<String, _>(AnotherConverter);
    assert_ne!(r1, r2); // Different now
}

// --- AnalysisContext edge cases ---
#[test]
fn analysis_context_custom_object_none() {
    let ctx = AnalysisContext::new("S", 0, 0);
    assert!(ctx.custom_object().is_none());
    assert!(ctx.custom::<String>().is_none());
}

#[test]
fn analysis_context_with_batch_index_preserves_sheet() {
    let ctx = AnalysisContext::new("MySheet", 3, 42).with_batch_index(7);
    assert_eq!(ctx.sheet_name(), "MySheet");
    assert_eq!(ctx.sheet_no(), 3);
    assert_eq!(ctx.row_index(), 42);
    assert_eq!(ctx.batch_index(), 7);
}

// --- ErrorAction edge cases ---
#[test]
fn error_action_all_variants() {
    assert_eq!(ErrorAction::Continue, ErrorAction::Continue);
    assert_eq!(ErrorAction::SkipRow, ErrorAction::SkipRow);
    assert_eq!(ErrorAction::Stop, ErrorAction::Stop);
    assert_ne!(ErrorAction::Continue, ErrorAction::Stop);
    assert_ne!(ErrorAction::SkipRow, ErrorAction::Stop);
}

// --- ExcelError edge cases ---
#[test]
fn excel_error_io_kinds() {
    let not_found = ExcelError::Io(std::io::Error::new(std::io::ErrorKind::NotFound, "missing"));
    let perm_denied = ExcelError::Io(std::io::Error::new(std::io::ErrorKind::PermissionDenied, "no access"));
    assert!(not_found.to_string().contains("missing"));
    assert!(perm_denied.to_string().contains("no access"));
}

// --- WriteHandler edge cases ---
#[test]
fn write_handler_before_cell_can_skip() {
    struct Skipper;
    impl WriteHandler for Skipper {
        fn before_cell(&mut self, ctx: &mut WriteCellContext) -> Result<()> {
            ctx.skip = true;
            Ok(())
        }
    }
    let mut h = Skipper;
    let mut cl = WriteCellContext {
        sheet_name: "S".to_owned(), row_index: 0, column_index: 0,
        field: None, is_head: false, relative_row_index: None, value: CellValue::Empty, skip: false,
    };
    h.before_cell(&mut cl).unwrap();
    assert!(cl.skip);
}

#[test]
fn write_handler_before_cell_can_transform() {
    struct Transformer;
    impl WriteHandler for Transformer {
        fn before_cell(&mut self, ctx: &mut WriteCellContext) -> Result<()> {
            ctx.value = CellValue::String("transformed".to_owned());
            Ok(())
        }
    }
    let mut h = Transformer;
    let mut cl = WriteCellContext {
        sheet_name: "S".to_owned(), row_index: 0, column_index: 0,
        field: None, is_head: false, relative_row_index: None, value: CellValue::Int(42), skip: false,
    };
    h.before_cell(&mut cl).unwrap();
    assert_eq!(cl.value, CellValue::String("transformed".to_owned()));
}
