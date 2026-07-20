//! End-to-end compatibility tests for the public facade.

use std::cell::RefCell;
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::net::TcpListener;
use std::path::PathBuf;
use std::rc::Rc;
use std::thread;

use chrono::NaiveDate;
use easyexcel::{
    AnalysisContext, AnchorType, BigInt, CellStyle, CellValue, ClientAnchorData, Converter,
    CoordinateData, EasyExcel, ExcelColor, ExcelColumn, ExcelError, ExcelFontScript, ExcelRow,
    ExcelUnderline, HorizontalAlignment, ImageData, ImageInputStream, InputStreamImageConverter,
    IntoExcelCell, LoopMergeProperty, OnceAbsoluteMergeProperty, PageReadListener,
    ReadConverterContext, ReadListener, Result, RichTextStringData, Url, UrlImageConverter,
    VerticalAlignment, WriteCellData, WriteConverterContext, WriteFont,
};
use tempfile::tempdir;
use zip::ZipArchive;

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

fn test_user(name: &str, age: u32) -> User {
    User {
        name: name.to_owned(),
        age: Some(age),
        registered_on: NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid test date"),
        transient: String::new(),
    }
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct ImageConverterRow {
    #[excel(name = "Primitive bytes", index = 0)]
    primitive_bytes: Vec<u8>,
    #[excel(name = "Boxed bytes", index = 1)]
    boxed_bytes: Box<[u8]>,
    #[excel(name = "Fixed bytes", index = 2)]
    fixed_bytes: [u8; 70],
    #[excel(name = "File", index = 3)]
    file: PathBuf,
    #[excel(name = "String file", index = 4, converter = easyexcel::StringImageConverter)]
    string_file: String,
}

#[derive(Debug, ExcelRow)]
struct StreamUrlImageRow {
    #[excel(name = "InputStream", index = 0, converter = InputStreamImageConverter)]
    stream: ImageInputStream<Cursor<Vec<u8>>>,
    #[excel(name = "URL", index = 1, converter = UrlImageConverter)]
    url: Url,
}

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct MultiImageRow {
    #[excel(name = "Images", index = 0)]
    cell: WriteCellData,
}

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct RichTextFacadeRow {
    #[excel(name = "Rich", index = 0)]
    value: RichTextStringData,
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

#[derive(Default)]
struct FormulaConverter;

impl Converter<String> for FormulaConverter {
    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<String> {
        Ok(context
            .formula()
            .map_or_else(String::new, |formula| formula.formula_value().to_owned()))
    }

    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<CellValue> {
        Ok(CellValue::Formula(context.value().clone()))
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
struct FormulaExpression {
    #[excel(name = "Formula", index = 0, converter = FormulaConverter)]
    formula: String,
}

#[derive(Debug, Clone, PartialEq, ExcelRow)]
struct CachedFormulaValue {
    #[excel(name = "Formula", index = 0)]
    value: f64,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct LargeInteger {
    #[excel(name = "整数", index = 0)]
    value: i64,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct ArbitraryInteger {
    #[excel(name = "BigInteger", index = 0)]
    value: BigInt,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
#[excel(column_width = 18, head_row_height = 24, content_row_height = 16)]
struct AnnotatedDimensions {
    #[excel(name = "姓名", index = 0, column_width = 30)]
    name: String,
    #[excel(name = "年龄", index = 1)]
    age: u32,
}

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
#[excel(
    head_style(
        horizontal_alignment = "center",
        fill_pattern = "solid",
        fill_foreground_color = 0x00ff_0000,
        border_bottom = "thin"
    ),
    content_style(wrapped = true),
    head_font_style(font_name = "Arial", font_height_in_points = 14, bold = true),
    content_font_style(italic = true),
    once_absolute_merge(
        first_row_index = 0,
        last_row_index = 0,
        first_column_index = 0,
        last_column_index = 1
    )
)]
struct AnnotatedStyles {
    #[excel(
        name = "姓名",
        index = 0,
        head_style(fill_pattern = "solid", fill_foreground_color = 0x0000_00ff),
        head_font_style(font_height_in_points = 20),
        content_loop_merge(each_row = 2, column_extend = 1)
    )]
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

fn serve_image_once(
    status: &str,
    body: Vec<u8>,
    declared_length: usize,
) -> Result<(Url, thread::JoinHandle<()>)> {
    let listener = TcpListener::bind("127.0.0.1:0")?;
    let address = listener.local_addr()?;
    let status = status.to_owned();
    let server = thread::spawn(move || {
        let (mut socket, _) = listener.accept().expect("accept image request");
        let mut request = [0_u8; 1024];
        let _ = socket.read(&mut request).expect("read image request");
        write!(
            socket,
            "HTTP/1.1 {status}\r\nContent-Length: {declared_length}\r\nConnection: close\r\n\r\n"
        )
        .expect("write image response head");
        socket.write_all(&body).expect("write image response body");
    });
    let url = Url::parse(&format!("http://{address}/logo.png"))
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    Ok((url, server))
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
fn stateful_csv_finishes_a_real_multi_batch_public_workflow() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("users.csv");
    let sheet = EasyExcel::writer_sheet::<User>("用户");
    let mut writer = EasyExcel::write::<User>(&path).with_bom(false).build();
    writer
        .write([test_user("张三", 30)], &sheet)?
        .write([test_user("李四", 31)], &sheet)?;
    writer.finish()?;
    writer.finish()?;
    let mut empty_writer = EasyExcel::write::<User>(directory.path().join("empty.csv")).build();
    empty_writer.finish()?;

    let rows = EasyExcel::read_sync::<User>(&path).do_read_sync()?;
    assert_eq!(rows, [test_user("张三", 30), test_user("李四", 31)]);
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
fn java_big_integer_fields_round_trip_without_precision_loss() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("big-integers.xlsx");
    let values = vec![
        ArbitraryInteger {
            value: BigInt::from(42),
        },
        ArbitraryInteger {
            value: "1234567890123456789012345678901234567890"
                .parse()
                .expect("valid big integer"),
        },
        ArbitraryInteger {
            value: "-987654321098765432109876543210987654321"
                .parse()
                .expect("valid big integer"),
        },
    ];

    EasyExcel::write::<ArbitraryInteger>(&path).do_write(values.clone())?;
    assert_eq!(
        EasyExcel::read_sync::<ArbitraryInteger>(&path).do_read_sync()?,
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
fn derive_uses_java_style_byte_array_and_file_image_converters() -> Result<()> {
    let directory = tempdir()?;
    let image_path = directory.path().join("source.png");
    let bytes = tiny_png();
    std::fs::write(&image_path, &bytes)?;
    let fixed_bytes: [u8; 70] = bytes.clone().try_into().expect("70-byte PNG fixture");
    let workbook_path = directory.path().join("image-converters.xlsx");

    EasyExcel::write::<ImageConverterRow>(&workbook_path).do_write([ImageConverterRow {
        primitive_bytes: bytes.clone(),
        boxed_bytes: bytes.clone().into_boxed_slice(),
        fixed_bytes,
        file: image_path.clone(),
        string_file: image_path.to_string_lossy().into_owned(),
    }])?;

    let mut archive = ZipArchive::new(File::open(&workbook_path)?)
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    let media_entries = (0..archive.len())
        .map(|index| {
            archive
                .by_index(index)
                .map(|entry| entry.name().to_owned())
                .map_err(|error| ExcelError::Format(error.to_string()))
        })
        .collect::<Result<Vec<_>>>()?
        .into_iter()
        .filter(|name| name.starts_with("xl/media/"))
        .count();
    assert_eq!(media_entries, 1);
    let mut drawing_xml = String::new();
    archive
        .by_name("xl/drawings/drawing1.xml")
        .map_err(|error| ExcelError::Format(error.to_string()))?
        .read_to_string(&mut drawing_xml)?;
    assert_eq!(drawing_xml.matches("<xdr:twoCellAnchor").count(), 5);
    Ok(())
}

#[test]
fn derive_uses_java_style_input_stream_and_url_image_converters() -> Result<()> {
    let bytes = tiny_png();
    let probe_stream = ImageInputStream::from(Cursor::new(bytes.clone()));
    assert_eq!(probe_stream.into_inner().into_inner(), bytes);
    let defaults = UrlImageConverter::default();
    assert_eq!(
        defaults.connect_timeout(),
        UrlImageConverter::DEFAULT_CONNECT_TIMEOUT
    );
    assert_eq!(
        defaults.read_timeout(),
        UrlImageConverter::DEFAULT_READ_TIMEOUT
    );
    let (url, server) = serve_image_once("200 OK", bytes.clone(), bytes.len())?;
    let directory = tempdir()?;
    let workbook_path = directory.path().join("stream-url-images.xlsx");

    EasyExcel::write::<StreamUrlImageRow>(&workbook_path).do_write([StreamUrlImageRow {
        stream: ImageInputStream::new(Cursor::new(bytes.clone())),
        url,
    }])?;
    server.join().expect("image server joins");

    let conversion = easyexcel::ConvertContext {
        sheet_name: "Images".to_owned(),
        row_index: 1,
        column_index: Some(1),
        field: "url",
        format: None,
    };
    let (url, server) = serve_image_once("404 Not Found", Vec::new(), 0)?;
    assert!(url.to_excel_cell(&conversion).is_err());
    server.join().expect("image server joins");
    let (url, server) = serve_image_once("200 OK", bytes.clone(), bytes.len() + 1)?;
    assert!(url.to_excel_cell(&conversion).is_err());
    server.join().expect("image server joins");
    let column = ExcelColumn::new("stream", "InputStream", Some(0), 0, None);
    let read_context = ReadConverterContext::new(None, &column, &conversion);
    assert!(
        Converter::<ImageInputStream<Cursor<Vec<u8>>>>::convert_to_rust_data(
            &InputStreamImageConverter,
            &read_context,
        )
        .is_err()
    );

    let mut archive = ZipArchive::new(File::open(&workbook_path)?)
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    let mut drawing_xml = String::new();
    archive
        .by_name("xl/drawings/drawing1.xml")
        .map_err(|error| ExcelError::Format(error.to_string()))?
        .read_to_string(&mut drawing_xml)?;
    assert_eq!(drawing_xml.matches("<xdr:twoCellAnchor").count(), 2);
    Ok(())
}

#[test]
fn public_facade_round_trips_scalar_write_cell_data_and_emits_multiple_images() -> Result<()> {
    let bytes = tiny_png();
    let second_anchor = ClientAnchorData::new()
        .coordinates(
            CoordinateData::new()
                .relative_first_column_index(1)
                .relative_last_column_index(1),
        )
        .left(3)
        .top(4)
        .right(5)
        .bottom(6)
        .anchor_type(AnchorType::DontMoveAndResize);
    let absolute_anchor = ClientAnchorData::new().coordinates(
        CoordinateData::new()
            .first_row_index(1)
            .first_column_index(1)
            .last_row_index(1)
            .last_column_index(1),
    );
    let row = MultiImageRow {
        cell: WriteCellData::new(CellValue::String("three images".to_owned())).image_data_list([
            ImageData::new(bytes.clone()),
            ImageData::new(bytes.clone()).anchor(second_anchor),
            ImageData::new(bytes).anchor(absolute_anchor),
        ]),
    };
    let directory = tempdir()?;
    let path = directory.path().join("multi-image.xlsx");
    EasyExcel::write::<MultiImageRow>(&path)
        .sheet("Images")
        .column_width(0, 18)
        .column_width(1, 12)
        .do_write([row])?;

    let rows = EasyExcel::read_sync::<MultiImageRow>(&path)
        .sheet("Images")
        .do_read_sync()?;
    assert_eq!(rows.len(), 1);
    assert_eq!(
        rows[0].cell.value(),
        &CellValue::String("three images".to_owned())
    );
    assert!(rows[0].cell.images().is_empty());

    let mut archive = ZipArchive::new(File::open(&path)?)
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    let mut drawing_xml = String::new();
    archive
        .by_name("xl/drawings/drawing1.xml")
        .map_err(|error| ExcelError::Format(error.to_string()))?
        .read_to_string(&mut drawing_xml)?;
    assert_eq!(drawing_xml.matches("<xdr:twoCellAnchor").count(), 3);
    assert_eq!(drawing_xml.matches("editAs=\"absolute\"").count(), 1);
    Ok(())
}

#[test]
fn public_facade_writes_rich_text_and_reads_its_plain_value() -> Result<()> {
    let rich = RichTextStringData::new("红色😀下标")
        .apply_font(
            WriteFont::new()
                .font_name("Aptos")
                .bold(true)
                .type_offset(ExcelFontScript::None),
        )
        .apply_font_range(
            0,
            2,
            WriteFont::new()
                .color(ExcelColor::Indexed(10))
                .underline(ExcelUnderline::Single),
        )
        .apply_font_range(
            2,
            4,
            WriteFont::new()
                .color(ExcelColor::Rgb(0x00_80_00))
                .type_offset(ExcelFontScript::Subscript),
        )
        .apply_font_range(
            0,
            1,
            WriteFont::new().type_offset(ExcelFontScript::Superscript),
        );
    let directory = tempdir()?;
    let path = directory.path().join("rich-text.xlsx");
    EasyExcel::write::<RichTextFacadeRow>(&path)
        .sheet("Rich")
        .do_write([RichTextFacadeRow {
            value: rich.clone(),
        }])?;

    for (name, value) in [
        (
            "outside.xlsx",
            RichTextStringData::new("a").apply_font_range(0, 2, WriteFont::new()),
        ),
        (
            "surrogate.xlsx",
            RichTextStringData::new("😀").apply_font_range(0, 1, WriteFont::new()),
        ),
    ] {
        assert!(
            EasyExcel::write::<RichTextFacadeRow>(directory.path().join(name))
                .sheet("Rich")
                .do_write([RichTextFacadeRow { value }])
                .is_err()
        );
    }

    let rows = EasyExcel::read_sync::<RichTextFacadeRow>(&path)
        .sheet("Rich")
        .do_read_sync()?;
    assert_eq!(rows[0].value.text_string(), rich.text_string());
    assert!(rows[0].value.interval_fonts().is_empty());
    let mut archive = ZipArchive::new(File::open(&path)?)
        .map_err(|error| ExcelError::Format(error.to_string()))?;
    let mut shared_strings = String::new();
    archive
        .by_name("xl/sharedStrings.xml")
        .map_err(|error| ExcelError::Format(error.to_string()))?
        .read_to_string(&mut shared_strings)?;
    assert!(shared_strings.contains("<r>"));
    assert!(shared_strings.contains('红'));
    assert!(shared_strings.contains('色'));
    assert!(shared_strings.contains("😀"));
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
fn formula_converter_receives_expression_while_scalar_receives_cached_value() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("formula.xlsx");
    let expected = vec![FormulaExpression {
        formula: "SUM(1,2)".to_owned(),
    }];
    EasyExcel::write::<FormulaExpression>(&path).do_write(expected.clone())?;

    assert_eq!(
        EasyExcel::read_sync::<FormulaExpression>(&path).do_read_sync()?,
        expected
    );
    assert_eq!(
        EasyExcel::read_sync::<CachedFormulaValue>(&path).do_read_sync()?,
        vec![CachedFormulaValue { value: 0.0 }]
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
fn derive_writes_java_style_cell_and_font_annotations() -> Result<()> {
    let metadata = AnnotatedStyles::write_metadata();
    assert!(metadata.head_style.is_some());
    assert!(metadata.content_style.is_some());
    assert!(metadata.head_font_style.is_some());
    assert!(metadata.content_font_style.is_some());
    assert_eq!(
        metadata.once_absolute_merge,
        Some(OnceAbsoluteMergeProperty::new(0, 0, 0, 1))
    );
    assert!(AnnotatedStyles::schema()[0].head_style.is_some());
    assert!(AnnotatedStyles::schema()[0].head_font_style.is_some());
    assert_eq!(
        AnnotatedStyles::schema()[0].loop_merge,
        Some(LoopMergeProperty::new(2, 1))
    );

    let directory = tempdir()?;
    EasyExcel::write::<AnnotatedStyles>(directory.path().join("annotated-styles.xlsx")).do_write([
        AnnotatedStyles {
            name: "Alice".to_owned(),
            age: 30,
        },
        AnnotatedStyles {
            name: "Bob".to_owned(),
            age: 31,
        },
    ])?;
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
