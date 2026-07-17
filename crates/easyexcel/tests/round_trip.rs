//! End-to-end compatibility tests for the public facade.

use std::cell::RefCell;
use std::rc::Rc;

use chrono::NaiveDate;
use easyexcel::{
    AnalysisContext, CellStyle, CellValue, Converter, EasyExcel, ExcelColumn, ExcelError, ExcelRow,
    HorizontalAlignment, PageReadListener, ReadConverterContext, ReadListener, Result,
    VerticalAlignment, WriteConverterContext,
};
use tempfile::tempdir;

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct User {
    #[excel(name = "姓名", index = 0)]
    name: String,
    #[excel(name = "年龄", index = 1)]
    age: Option<u32>,
    #[excel(name = "注册日期", index = 2, format = "%Y-%m-%d")]
    registered_on: NaiveDate,
    #[excel(ignore)]
    transient: String,
}

#[derive(Default)]
struct NameConverter;

impl Converter<String> for NameConverter {
    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<String> {
        Ok(context
            .cell()
            .map_or_else(String::new, CellValue::as_text)
            .strip_prefix("excel:")
            .unwrap_or_default()
            .to_owned())
    }

    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<CellValue> {
        Ok(CellValue::String(format!("excel:{}", context.value())))
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct ConvertedName {
    #[excel(name = "姓名", index = 0, converter = NameConverter)]
    name: String,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct RawName {
    #[excel(name = "姓名", index = 0)]
    name: String,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct LargeInteger {
    #[excel(name = "整数", index = 0)]
    value: i64,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
#[excel(column_width = 18, head_row_height = 24, content_row_height = 16)]
struct AnnotatedDimensions {
    #[excel(name = "姓名", index = 0, column_width = 30)]
    name: String,
    #[excel(name = "年龄", index = 1)]
    age: u32,
}

struct EveryPublicCell;

impl ExcelRow for EveryPublicCell {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[
            ExcelColumn::new("empty", "Empty", Some(0), 0, None),
            ExcelColumn::new("string", "String", Some(1), 0, None),
            ExcelColumn::new("error", "Error", Some(2), 0, None),
            ExcelColumn::new("boolean", "Boolean", Some(3), 0, None),
            ExcelColumn::new("integer", "Integer", Some(4), 0, None),
            ExcelColumn::new("float", "Float", Some(5), 0, None),
            ExcelColumn::new("date", "Date", Some(6), 0, Some("%d/%m/%Y")),
            ExcelColumn::new(
                "datetime",
                "DateTime",
                Some(7),
                0,
                Some("%Y-%m-%d %H:%M:%S"),
            ),
            ExcelColumn::new("large", "Large", Some(8), 0, None),
            ExcelColumn::new("formula", "Formula", Some(9), 0, None),
            ExcelColumn::new("link", "Link", Some(10), 0, None),
            ExcelColumn::new("comment", "Comment", Some(11), 0, None),
            ExcelColumn::new("image", "Image", Some(12), 0, None),
        ];
        COLUMNS
    }

    fn from_row(_row: &easyexcel::RowData) -> Result<Self> {
        Err(ExcelError::Unsupported("write-only test row".to_owned()))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
        Ok(vec![
            CellValue::Empty,
            CellValue::String("text".to_owned()),
            CellValue::Error("#DIV/0!".to_owned()),
            CellValue::Bool(true),
            CellValue::Int(-12),
            CellValue::Float(1.25),
            CellValue::Date(date),
            CellValue::DateTime(date.and_hms_opt(12, 34, 56).expect("valid time")),
            CellValue::Int(i64::MAX),
            CellValue::Formula("SUM(E2:F2)".to_owned()),
            CellValue::Hyperlink {
                url: "https://www.rust-lang.org".to_owned(),
                text: "Rust".to_owned(),
            },
            CellValue::Comment {
                value: Box::new(CellValue::String("annotated".to_owned())),
                text: "cell note".to_owned(),
            },
            CellValue::Image(tiny_png()),
        ])
    }
}

fn tiny_png() -> Vec<u8> {
    vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f,
        0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0xf0, 0x1f, 0x00, 0x05, 0x00, 0x01, 0xff, 0x89, 0x99, 0x3d, 0x1d, 0x00, 0x00,
        0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ]
}

#[test]
fn writes_and_reads_typed_rows_with_java_style_builders() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("users.xlsx");
    let users = vec![
        User {
            name: "张三".to_owned(),
            age: Some(30),
            registered_on: NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date"),
            transient: String::new(),
        },
        User {
            name: "李四".to_owned(),
            age: None,
            registered_on: NaiveDate::from_ymd_opt(2025, 1, 2).expect("valid date"),
            transient: String::new(),
        },
    ];

    EasyExcel::write::<User>(&path)
        .sheet("用户")
        .freeze_head(true)
        .constant_memory(true)
        .do_write_iter(users.clone())?;

    let actual = EasyExcel::read_sync::<User>(&path)
        .sheet("用户")
        .do_read_sync()?;
    assert_eq!(actual, users);
    Ok(())
}

#[test]
fn integers_beyond_excels_exact_number_range_round_trip_as_text() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("large-integers.xlsx");
    let values = vec![
        LargeInteger { value: 42 },
        LargeInteger { value: i64::MAX },
        LargeInteger { value: i64::MIN },
        LargeInteger { value: 1 },
        LargeInteger { value: 2 },
        LargeInteger { value: 3 },
        LargeInteger { value: 4 },
    ];
    EasyExcel::write::<LargeInteger>(&path)
        .content_styles([
            CellStyle::new()
                .italic(true)
                .font_color(0x11_22_33)
                .background_color(0xEE_DD_CC)
                .horizontal_alignment(HorizontalAlignment::General)
                .vertical_alignment(VerticalAlignment::Top)
                .wrap_text(true)
                .number_format("0"),
            CellStyle::new()
                .horizontal_alignment(HorizontalAlignment::Left)
                .vertical_alignment(VerticalAlignment::Center)
                .bold(true),
            CellStyle::new()
                .horizontal_alignment(HorizontalAlignment::Center)
                .vertical_alignment(VerticalAlignment::Bottom),
            CellStyle::new()
                .horizontal_alignment(HorizontalAlignment::Right)
                .vertical_alignment(VerticalAlignment::Justify),
            CellStyle::new()
                .horizontal_alignment(HorizontalAlignment::Fill)
                .vertical_alignment(VerticalAlignment::Distributed),
            CellStyle::new().horizontal_alignment(HorizontalAlignment::Justify),
            CellStyle::new().horizontal_alignment(HorizontalAlignment::CenterAcross),
        ])
        .do_write(values.clone())?;
    assert_eq!(
        EasyExcel::read_sync::<LargeInteger>(&path).do_read_sync()?,
        values
    );
    Ok(())
}

#[test]
fn public_writer_accepts_every_supported_cell_variant() -> Result<()> {
    let directory = tempdir()?;
    EasyExcel::write::<EveryPublicCell>(directory.path().join("every-cell.xlsx"))
        .do_write([EveryPublicCell])?;
    Ok(())
}

#[test]
fn derive_selected_converter_transforms_read_and_write_values() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("converted.xlsx");
    let expected = vec![ConvertedName {
        name: "alice".to_owned(),
    }];
    EasyExcel::write::<ConvertedName>(&path).do_write(expected.clone())?;
    assert_eq!(
        EasyExcel::read_sync::<RawName>(&path).do_read_sync()?,
        vec![RawName {
            name: "excel:alice".to_owned()
        }]
    );
    assert_eq!(
        EasyExcel::read_sync::<ConvertedName>(&path).do_read_sync()?,
        expected
    );
    Ok(())
}

#[test]
fn derive_exposes_java_style_dimension_annotations() -> Result<()> {
    let metadata = AnnotatedDimensions::write_metadata();
    assert_eq!(AnnotatedDimensions::schema()[0].column_width, Some(30));
    assert_eq!(AnnotatedDimensions::schema()[1].column_width, None);
    assert_eq!(metadata.column_width, Some(18));
    assert_eq!(metadata.head_row_height, Some(24));
    assert_eq!(metadata.content_row_height, Some(16));

    let directory = tempdir()?;
    EasyExcel::write::<AnnotatedDimensions>(directory.path().join("dimensions.xlsx"))
        .column_width(1, 40)
        .do_write([AnnotatedDimensions {
            name: "Alice".to_owned(),
            age: 30,
        }])?;
    Ok(())
}

#[test]
fn page_listener_receives_batches_and_contexts() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("users.xlsx");
    let users = (0..4)
        .map(|age| User {
            name: format!("user-{age}"),
            age: Some(age),
            registered_on: NaiveDate::from_ymd_opt(2026, 1, 1).expect("valid date"),
            transient: String::new(),
        })
        .collect::<Vec<_>>();
    EasyExcel::write::<User>(&path).do_write(users)?;

    let batches = Rc::new(RefCell::new(Vec::new()));
    let captured = Rc::clone(&batches);
    let listener = PageReadListener::new(2, move |rows: Vec<User>, context| {
        captured
            .borrow_mut()
            .push((rows.len(), context.batch_index()));
        Ok(())
    });
    EasyExcel::read::<User, _>(&path, listener).do_read()?;

    assert_eq!(&*batches.borrow(), &[(2, 0), (2, 1)]);
    Ok(())
}

struct StopListener;

impl ReadListener<User> for StopListener {
    fn invoke(&mut self, _data: User, _context: &AnalysisContext) -> Result<()> {
        panic!("has_next prevents invocation")
    }

    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        false
    }
}

#[test]
fn listener_can_stop_before_data_rows() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("users.xlsx");
    let user = User {
        name: "stop".to_owned(),
        age: Some(1),
        registered_on: NaiveDate::from_ymd_opt(2026, 1, 1).expect("valid date"),
        transient: String::new(),
    };
    EasyExcel::write::<User>(&path).do_write([user])?;
    EasyExcel::read::<User, _>(&path, StopListener)
        .sheet(0_usize)
        .head_row_number(1)
        .ignore_empty_row(false)
        .do_read()
}
