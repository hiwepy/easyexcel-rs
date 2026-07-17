use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{self, Cursor, Read, Write};
use std::net::TcpListener;
use std::sync::Arc;
use std::thread;
use std::time::Duration;

use chrono::{NaiveDate, NaiveDateTime};

use super::*;

#[test]
fn csv_charset_accepts_java_style_names_and_has_a_utf8_default() {
    assert_eq!(CsvCharset::default(), CsvCharset::utf8());
    assert_eq!(CsvCharset::default().name(), "UTF-8");
    assert_eq!(CsvCharset::from("GBK").name(), "GBK");
    assert_eq!(CsvCharset::from("UTF-16BE".to_owned()).name(), "UTF-16BE");
}

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
        (
            CellValue::Decimal(BigDecimal::from(1)),
            CellDataType::Number,
        ),
        (CellValue::Date(date), CellDataType::Date),
        (CellValue::DateTime(datetime), CellDataType::Date),
        (CellValue::Error("#N/A".to_owned()), CellDataType::Error),
        (CellValue::Formula("1+1".to_owned()), CellDataType::Formula),
        (
            CellValue::Hyperlink {
                url: "https://example.com".to_owned(),
                text: "link".to_owned(),
            },
            CellDataType::String,
        ),
        (
            CellValue::Comment {
                value: Box::new(CellValue::Bool(false)),
                text: "note".to_owned(),
            },
            CellDataType::Boolean,
        ),
        (CellValue::Image(vec![1]), CellDataType::Image),
    ];
    for (cell, expected) in cases {
        assert_eq!(cell.data_type(), expected);
    }
    assert_ne!(CellDataType::DirectString, CellDataType::RichTextString);
}

#[test]
fn image_converters_match_java_byte_array_and_file_write_semantics() -> Result<()> {
    let conversion = context(None);
    let bytes = vec![0x89, b'P', b'N', b'G'];
    let image = CellValue::Image(bytes.clone());

    assert_eq!(
        Vec::<u8>::from_excel_cell(Some(&image), &conversion)?,
        bytes
    );
    assert!(Vec::<u8>::from_excel_cell(Some(&CellValue::Empty), &conversion).is_err());
    assert_eq!(bytes.to_excel_cell(&conversion)?, image);

    let boxed = bytes.clone().into_boxed_slice();
    assert_eq!(
        Box::<[u8]>::from_excel_cell(Some(&image), &conversion)?,
        boxed
    );
    assert_eq!(boxed.to_excel_cell(&conversion)?, image);

    let fixed = <[u8; 4]>::from_excel_cell(Some(&image), &conversion)?;
    assert_eq!(fixed, [0x89, b'P', b'N', b'G']);
    assert_eq!(fixed.to_excel_cell(&conversion)?, image);
    let short_image = CellValue::Image(vec![1, 2, 3]);
    assert_eq!(
        <[u8; 3]>::from_excel_cell(Some(&short_image), &conversion)?,
        [1, 2, 3]
    );
    assert!(<[u8; 3]>::from_excel_cell(Some(&image), &conversion).is_err());
    assert!(<[u8; 4]>::from_excel_cell(Some(&short_image), &conversion).is_err());
    assert!(<[u8; 4]>::from_excel_cell(Some(&CellValue::Empty), &conversion).is_err());

    let path = PathBuf::from_excel_cell(
        Some(&CellValue::String("images/logo.png".to_owned())),
        &conversion,
    )?;
    assert_eq!(path, PathBuf::from("images/logo.png"));

    let directory = tempfile::tempdir()?;
    let image_path = directory.path().join("logo.png");
    std::fs::write(&image_path, &bytes)?;
    assert_eq!(image_path.to_excel_cell(&conversion)?, image);
    assert!(
        directory
            .path()
            .join("missing.png")
            .to_excel_cell(&conversion)
            .is_err()
    );

    let column = ExcelColumn::new("image", "Image", Some(0), 0, None);
    let path = image_path.to_string_lossy().into_owned();
    let write_context = WriteConverterContext::new(&path, &column, &conversion);
    assert_eq!(
        StringImageConverter.convert_to_excel_data(&write_context)?,
        image
    );
    let missing = directory
        .path()
        .join("missing.png")
        .to_string_lossy()
        .into_owned();
    let write_context = WriteConverterContext::new(&missing, &column, &conversion);
    assert!(
        StringImageConverter
            .convert_to_excel_data(&write_context)
            .is_err()
    );
    Ok(())
}

struct FailingImageReader;

impl Read for FailingImageReader {
    fn read(&mut self, _buffer: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::other("injected image stream failure"))
    }
}

fn serve_image_once(
    status: &str,
    body: Vec<u8>,
    declared_length: usize,
) -> (Url, thread::JoinHandle<()>) {
    let listener = TcpListener::bind("127.0.0.1:0").expect("bind test image server");
    let address = listener.local_addr().expect("test image server address");
    let status = status.to_owned();
    let handle = thread::spawn(move || {
        let (mut stream, _) = listener.accept().expect("accept image request");
        let mut request = [0_u8; 1024];
        let _ = stream.read(&mut request).expect("read image request");
        write!(
            stream,
            "HTTP/1.1 {status}\r\nContent-Length: {declared_length}\r\nConnection: close\r\n\r\n"
        )
        .expect("write image response head");
        stream.write_all(&body).expect("write image response body");
    });
    (
        Url::parse(&format!("http://{address}/logo.png")).expect("valid local image URL"),
        handle,
    )
}

#[test]
fn input_stream_and_url_converters_match_java_ownership_and_timeout_semantics() -> Result<()> {
    let conversion = context(None);
    let bytes = vec![0x89, b'P', b'N', b'G'];
    let stream = ImageInputStream::from(Cursor::new(bytes.clone()));
    assert_eq!(
        stream.to_excel_cell(&conversion)?,
        CellValue::Image(bytes.clone())
    );
    assert_eq!(
        stream.to_excel_cell(&conversion)?,
        CellValue::Image(Vec::new())
    );
    assert_eq!(stream.into_inner().position(), bytes.len() as u64);
    assert!(
        ImageInputStream::new(FailingImageReader)
            .to_excel_cell(&conversion)
            .is_err()
    );
    let converter_stream = ImageInputStream::new(Cursor::new(bytes.clone()));
    let stream_column = ExcelColumn::new("stream", "Stream", Some(0), 0, None);
    let stream_context = WriteConverterContext::new(&converter_stream, &stream_column, &conversion);
    assert_eq!(
        InputStreamImageConverter.convert_to_excel_data(&stream_context)?,
        CellValue::Image(bytes.clone())
    );

    let defaults = UrlImageConverter::default();
    assert_eq!(
        defaults.connect_timeout(),
        UrlImageConverter::DEFAULT_CONNECT_TIMEOUT
    );
    assert_eq!(
        defaults.read_timeout(),
        UrlImageConverter::DEFAULT_READ_TIMEOUT
    );
    let custom = UrlImageConverter::new(Duration::from_secs(2), Duration::from_secs(3));
    assert_eq!(custom.connect_timeout(), Duration::from_secs(2));
    assert_eq!(custom.read_timeout(), Duration::from_secs(3));
    let (url, server) = serve_image_once("200 OK", bytes.clone(), bytes.len());
    assert_eq!(
        url.to_excel_cell(&conversion)?,
        CellValue::Image(bytes.clone())
    );
    server.join().expect("image server joins");

    let (url, server) = serve_image_once("200 OK", bytes.clone(), bytes.len());
    let column = ExcelColumn::new("image", "Image", Some(0), 0, None);
    let write_context = WriteConverterContext::new(&url, &column, &conversion);
    assert_eq!(
        custom.convert_to_excel_data(&write_context)?,
        CellValue::Image(bytes.clone())
    );
    server.join().expect("image server joins");

    let (url, server) = serve_image_once("404 Not Found", Vec::new(), 0);
    assert!(url.to_excel_cell(&conversion).is_err());
    server.join().expect("image server joins");

    let (url, server) = serve_image_once("200 OK", bytes.clone(), bytes.len() + 1);
    assert!(url.to_excel_cell(&conversion).is_err());
    server.join().expect("image server joins");
    Ok(())
}

#[test]
fn write_cell_image_data_matches_java_coordinate_anchor_and_list_model() -> Result<()> {
    let coordinates = CoordinateData::new()
        .first_row_index(3)
        .first_column_index(4)
        .last_row_index(5)
        .last_column_index(6)
        .relative_first_row_index(-1)
        .relative_first_column_index(-2)
        .relative_last_row_index(2)
        .relative_last_column_index(3);
    assert_eq!(coordinates.get_first_row_index(), Some(3));
    assert_eq!(coordinates.get_first_column_index(), Some(4));
    assert_eq!(coordinates.get_last_row_index(), Some(5));
    assert_eq!(coordinates.get_last_column_index(), Some(6));
    assert_eq!(coordinates.get_relative_first_row_index(), Some(-1));
    assert_eq!(coordinates.get_relative_first_column_index(), Some(-2));
    assert_eq!(coordinates.get_relative_last_row_index(), Some(2));
    assert_eq!(coordinates.get_relative_last_column_index(), Some(3));
    assert_eq!(CoordinateData::default(), CoordinateData::new());

    let anchor = ClientAnchorData::new()
        .coordinates(coordinates)
        .top(1)
        .right(2)
        .bottom(3)
        .left(4)
        .anchor_type(AnchorType::DontMoveAndResize);
    assert_eq!(anchor.get_coordinates(), coordinates);
    assert_eq!(anchor.get_top(), Some(1));
    assert_eq!(anchor.get_right(), Some(2));
    assert_eq!(anchor.get_bottom(), Some(3));
    assert_eq!(anchor.get_left(), Some(4));
    assert_eq!(
        anchor.get_anchor_type(),
        Some(AnchorType::DontMoveAndResize)
    );
    assert_eq!(ClientAnchorData::default(), ClientAnchorData::new());
    assert_eq!(AnchorType::default(), AnchorType::MoveAndResize);

    let image_types = [
        ImageType::Emf,
        ImageType::Wmf,
        ImageType::Pict,
        ImageType::Jpeg,
        ImageType::Png,
        ImageType::Dib,
    ];
    assert_eq!(image_types.len(), 6);
    let first = ImageData::new([1, 2, 3])
        .image_type(ImageType::Png)
        .anchor(anchor);
    assert_eq!(first.image(), &[1, 2, 3]);
    assert_eq!(first.get_image_type(), Some(ImageType::Png));
    assert_eq!(first.get_anchor(), anchor);
    assert_eq!(ImageData::default().image(), &[]);

    let second = ImageData::new(vec![4, 5, 6]);
    let data = WriteCellData::new(CellValue::String("caption".to_owned()))
        .image(first.clone())
        .image(second.clone());
    assert_eq!(data.value(), &CellValue::String("caption".to_owned()));
    assert_eq!(data.images(), &[first.clone(), second.clone()]);
    let replaced = data.clone().image_data_list([second.clone()]);
    assert_eq!(replaced.images(), &[second]);
    assert_eq!(
        WriteCellData::from_image(vec![7, 8]).images()[0].image(),
        &[7, 8]
    );
    let conversion = context(None);
    let converted = data.to_excel_cell(&conversion)?;
    assert_eq!(converted.as_text(), "caption");
    assert_eq!(converted.data_type(), CellDataType::String);
    assert!(!converted.is_empty());
    assert_eq!(
        WriteCellData::from_excel_cell(Some(&CellValue::Bool(true)), &conversion)?.value(),
        &CellValue::Bool(true)
    );
    assert_eq!(
        WriteCellData::from_excel_cell(None, &conversion)?.value(),
        &CellValue::Empty
    );
    Ok(())
}

#[test]
fn row_data_resolves_index_before_header_name() {
    let explicit = ExcelColumn::new("first", "Header", Some(1), 3, Some("0")).with_column_width(24);
    let named = ExcelColumn::new("second", "Header", None, i32::MAX, None);
    let missing = ExcelColumn::new("missing", "Missing", None, i32::MAX, None);
    let headers = Arc::new(HashMap::from([("Header".to_owned(), 0)]));
    let row = RowData::new(
        "Users",
        7,
        vec![CellValue::String("name".to_owned()), CellValue::Int(9)],
        headers,
    )
    .with_formulas(HashMap::from([
        (0, FormulaData::new("LOWER(A1)")),
        (1, FormulaData::new("1+8")),
    ]));

    assert_eq!(row.sheet_name(), "Users");
    assert_eq!(row.row_index(), 7);
    assert_eq!(row.cell(&explicit), Some(&CellValue::Int(9)));
    assert_eq!(
        row.cell(&named),
        Some(&CellValue::String("name".to_owned()))
    );
    assert_eq!(row.cell(&missing), None);
    assert_eq!(
        row.formula(&explicit).map(FormulaData::formula_value),
        Some("1+8")
    );
    assert_eq!(
        row.formula(&named).map(FormulaData::formula_value),
        Some("LOWER(A1)")
    );
    assert_eq!(row.formula(&missing), None);
    assert_eq!(FormulaData::default().formula_value(), "");
    assert_eq!(row.convert_context(&explicit).column_index, Some(1));
    assert_eq!(row.convert_context(&named).column_index, Some(0));
    assert_eq!(row.convert_context(&missing).column_index, None);
    assert_eq!(explicit.column_width, Some(24));
    assert_eq!(named.column_width, None);
    assert_eq!(ExcelWriteMetadata::default(), ExcelWriteMetadata::new());
    assert_eq!(ExcelColor::java_or_rgb(10), ExcelColor::Indexed(10));
    assert_eq!(
        ExcelColor::java_or_rgb(0x00ff_0000),
        ExcelColor::Rgb(0x00ff_0000)
    );
    let cell_style = ExcelCellStyle {
        horizontal_alignment: Some(ExcelHorizontalAlignment::Center),
        fill_pattern: Some(ExcelFillPattern::Solid),
        fill_foreground_color: Some(ExcelColor::Rgb(0x00ff_0000)),
        ..ExcelCellStyle::new()
    };
    let font_style = ExcelFontStyle {
        font_name: Some("Arial"),
        bold: Some(true),
        ..ExcelFontStyle::new()
    };
    let styled = ExcelColumn::new("styled", "Styled", None, 0, None)
        .with_head_style(cell_style)
        .with_content_style(cell_style)
        .with_head_font_style(font_style)
        .with_content_font_style(font_style);
    assert_eq!(styled.head_style, Some(cell_style));
    assert_eq!(styled.content_style, Some(cell_style));
    assert_eq!(styled.head_font_style, Some(font_style));
    assert_eq!(styled.content_font_style, Some(font_style));
    assert_eq!(
        ExcelWriteMetadata::new()
            .column_width(18)
            .head_row_height(20)
            .content_row_height(16)
            .head_style(cell_style)
            .content_style(cell_style)
            .head_font_style(font_style)
            .content_font_style(font_style),
        ExcelWriteMetadata {
            column_width: Some(18),
            head_row_height: Some(20),
            content_row_height: Some(16),
            head_style: Some(cell_style),
            content_style: Some(cell_style),
            head_font_style: Some(font_style),
            content_font_style: Some(font_style),
        }
    );
}

#[test]
fn dynamic_rows_match_java_no_model_return_modes() -> Result<()> {
    let headers = Arc::new(HashMap::from([("Tail".to_owned(), 4)]));
    let cells = vec![
        CellValue::String("value".to_owned()),
        CellValue::Empty,
        CellValue::Error("#N/A".to_owned()),
        CellValue::Empty,
    ];
    let row = RowData::new("Dynamic", 7, cells.clone(), Arc::clone(&headers))
        .with_formulas(HashMap::from([(2, FormulaData::new("NA()"))]))
        .with_display_values(HashMap::from([(2, "#N/A display".to_owned())]))
        .with_present_columns(HashSet::from([0, 2, 3]));

    assert_eq!(ReadDefaultReturn::default(), ReadDefaultReturn::String);
    let strings = DynamicRow::from_row(&row)?;
    assert_eq!(strings.values().len(), 5);
    assert_eq!(
        strings.get(0),
        Some(&DynamicValue::String("value".to_owned()))
    );
    assert_eq!(strings.get(1), Some(&DynamicValue::Null));
    assert_eq!(
        strings.get(2),
        Some(&DynamicValue::String("#N/A display".to_owned()))
    );
    assert_eq!(strings.get(3), Some(&DynamicValue::String(String::new())));
    assert_eq!(strings.get(4), Some(&DynamicValue::Null));

    let actual = DynamicRow::from_row(
        &RowData::new("Dynamic", 7, cells.clone(), Arc::clone(&headers))
            .with_present_columns(HashSet::from([0, 2, 3]))
            .with_read_default_return(ReadDefaultReturn::ActualData),
    )?;
    assert_eq!(
        actual.get(0),
        Some(&DynamicValue::ActualData(CellValue::String(
            "value".to_owned()
        )))
    );
    assert_eq!(
        actual.get(2),
        Some(&DynamicValue::ActualData(CellValue::String(
            "#N/A".to_owned()
        )))
    );
    assert_eq!(
        actual.get(3),
        Some(&DynamicValue::ActualData(CellValue::String(String::new())))
    );

    let cell_data_row = DynamicRow::from_row(
        &RowData::new("Dynamic", 7, cells, headers)
            .with_formulas(HashMap::from([(2, FormulaData::new("NA()"))]))
            .with_present_columns(HashSet::from([0, 2, 3]))
            .with_read_default_return(ReadDefaultReturn::ReadCellData),
    )?;
    let DynamicValue::ReadCellData(cell_data) = cell_data_row.get(2).expect("column 2") else {
        panic!("expected read cell data");
    };
    assert_eq!(cell_data.row_index(), 7);
    assert_eq!(cell_data.column_index(), 2);
    assert_eq!(cell_data.raw_value(), &CellValue::Error("#N/A".to_owned()));
    assert_eq!(cell_data.data(), &CellValue::String("#N/A".to_owned()));
    assert_eq!(cell_data.display_value(), "#N/A");
    assert_eq!(
        cell_data.formula().map(FormulaData::formula_value),
        Some("NA()")
    );
    let DynamicValue::ReadCellData(empty_data) = cell_data_row.get(3).expect("column 3") else {
        panic!("expected empty read cell data");
    };
    assert_eq!(empty_data.raw_value(), &CellValue::Empty);
    assert_eq!(empty_data.data(), &CellValue::String(String::new()));
    assert_eq!(empty_data.formula(), None);
    Ok(())
}

#[test]
fn dynamic_rows_are_ordered_and_can_be_written_back() -> Result<()> {
    let read_cell_data = ReadCellData::new(
        1,
        4,
        CellValue::Int(9),
        CellValue::Int(9),
        "9".to_owned(),
        None,
    );
    let row = DynamicRow::new(BTreeMap::from([
        (0, DynamicValue::String("first".to_owned())),
        (1, DynamicValue::Null),
        (2, DynamicValue::ActualData(CellValue::Bool(true))),
        (4, DynamicValue::ReadCellData(read_cell_data)),
    ]));
    assert!(DynamicRow::schema().is_empty());
    assert_eq!(
        row.clone()
            .into_values()
            .keys()
            .copied()
            .collect::<Vec<_>>(),
        vec![0, 1, 2, 4]
    );
    assert_eq!(
        row.to_row()?,
        vec![
            CellValue::String("first".to_owned()),
            CellValue::Empty,
            CellValue::Bool(true),
            CellValue::Empty,
            CellValue::Int(9),
        ]
    );
    assert!(DynamicRow::default().to_row()?.is_empty());
    let overflow = DynamicRow::new(BTreeMap::from([(usize::MAX, DynamicValue::Null)]));
    assert!(matches!(overflow.to_row(), Err(ExcelError::Format(_))));
    Ok(())
}

#[test]
fn strings_and_booleans_convert_in_both_directions() -> Result<()> {
    let context = context(None);
    assert_eq!(String::from_excel_cell(None, &context)?, "");
    assert_eq!(
        String::from_excel_cell(Some(&CellValue::Int(5)), &context)?,
        "5"
    );
    assert_eq!(
        "text".to_owned().to_excel_cell(&context)?,
        CellValue::String("text".to_owned())
    );
    assert_eq!(
        "borrowed".to_excel_cell(&context)?,
        CellValue::String("borrowed".to_owned())
    );

    for (value, expected) in [
        (CellValue::Bool(true), true),
        (CellValue::Int(1), true),
        (CellValue::Int(0), false),
        (CellValue::Float(0.5), true),
        (CellValue::Float(0.0), false),
        (
            CellValue::Decimal("0.5".parse().expect("valid decimal")),
            true,
        ),
        (
            CellValue::Decimal("0".parse().expect("valid decimal")),
            false,
        ),
        (CellValue::String("TRUE".to_owned()), true),
        (CellValue::String("1".to_owned()), true),
        (CellValue::String("false".to_owned()), false),
        (CellValue::String("0".to_owned()), false),
    ] {
        assert_eq!(bool::from_excel_cell(Some(&value), &context)?, expected);
    }
    assert!(bool::from_excel_cell(None, &context).is_err());
    assert_eq!(true.to_excel_cell(&context)?, CellValue::Bool(true));
    Ok(())
}

macro_rules! assert_integer_type {
    ($ty:ty, $value:expr) => {{
        let context = context(None);
        let parsed = <$ty>::from_excel_cell(Some(&CellValue::Int($value)), &context)?;
        assert_eq!(parsed.to_string(), $value.to_string());
        assert_eq!(parsed.to_excel_cell(&context)?, CellValue::Int($value));
        assert_eq!(
            <$ty>::from_excel_cell(Some(&CellValue::Float(7.0)), &context)?.to_string(),
            $value.to_string()
        );
        assert_eq!(
            <$ty>::from_excel_cell(Some(&CellValue::String($value.to_string())), &context)?
                .to_string(),
            $value.to_string()
        );
        assert_eq!(
            <$ty>::from_excel_cell(
                Some(&CellValue::Decimal(
                    $value.to_string().parse().expect("valid decimal"),
                )),
                &context,
            )?
            .to_string(),
            $value.to_string()
        );
        assert!(<$ty>::from_excel_cell(Some(&CellValue::Float(1.5)), &context).is_err());
        assert!(<$ty>::from_excel_cell(Some(&CellValue::Bool(true)), &context).is_err());
        assert!(
            <$ty>::from_excel_cell(Some(&CellValue::String("bad".to_owned())), &context).is_err()
        );
        assert!(<$ty>::from_excel_cell(None, &context).is_err());
    }};
}

#[test]
fn every_integer_type_is_supported_and_edge_paths_are_checked() -> Result<()> {
    assert_integer_type!(i8, 7);
    assert_integer_type!(i16, 7);
    assert_integer_type!(i32, 7);
    assert_integer_type!(i64, 7);
    assert_integer_type!(isize, 7);
    assert_integer_type!(u8, 7);
    assert_integer_type!(u16, 7);
    assert_integer_type!(u32, 7);
    assert_integer_type!(u64, 7);
    assert_integer_type!(usize, 7);

    let context = context(None);
    assert_eq!(
        i32::from_excel_cell(Some(&CellValue::Float(8.0)), &context)?,
        8
    );
    assert_eq!(
        i32::from_excel_cell(Some(&CellValue::String("9".to_owned())), &context)?,
        9
    );
    assert!(i32::from_excel_cell(Some(&CellValue::Float(8.5)), &context).is_err());
    assert!(
        i32::from_excel_cell(
            Some(&CellValue::Decimal("8.5".parse().expect("valid decimal"),)),
            &context,
        )
        .is_err()
    );
    assert!(i32::from_excel_cell(Some(&CellValue::Bool(true)), &context).is_err());
    assert!(i32::from_excel_cell(Some(&CellValue::String("bad".to_owned())), &context).is_err());
    assert!(u8::from_excel_cell(Some(&CellValue::Int(300)), &context).is_err());
    assert!(i32::from_excel_cell(None, &context).is_err());
    assert_eq!(
        u64::MAX.to_excel_cell(&context)?,
        CellValue::String(u64::MAX.to_string())
    );
    Ok(())
}

#[test]
fn big_integer_matches_java_boolean_number_and_string_converters() -> Result<()> {
    let context = context(None);
    let huge: BigInt = "123456789012345678901234567890"
        .parse()
        .expect("valid big integer");

    assert_eq!(
        BigInt::from_excel_cell(Some(&CellValue::Bool(true)), &context)?,
        BigInt::from(1)
    );
    assert_eq!(
        BigInt::from_excel_cell(Some(&CellValue::Bool(false)), &context)?,
        BigInt::from(0)
    );
    assert_eq!(
        BigInt::from_excel_cell(Some(&CellValue::Int(-12)), &context)?,
        BigInt::from(-12)
    );
    assert_eq!(
        BigInt::from_excel_cell(Some(&CellValue::Float(42.9)), &context)?,
        BigInt::from(42)
    );
    assert_eq!(
        BigInt::from_excel_cell(
            Some(&CellValue::Decimal(
                "-99.75".parse().expect("valid decimal"),
            )),
            &context,
        )?,
        BigInt::from(-99)
    );
    assert_eq!(
        BigInt::from_excel_cell(
            Some(&CellValue::String(
                "123456789012345678901234567890.8".to_owned(),
            )),
            &context,
        )?,
        huge
    );
    assert!(BigInt::from_excel_cell(Some(&CellValue::Float(f64::NAN)), &context).is_err());
    assert!(BigInt::from_excel_cell(Some(&CellValue::String("bad".to_owned())), &context).is_err());
    assert!(
        BigInt::from_excel_cell(
            Some(&CellValue::Date(
                NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date")
            )),
            &context
        )
        .is_err()
    );
    assert!(BigInt::from_excel_cell(None, &context).is_err());
    assert_eq!(BigInt::from(7).to_excel_cell(&context)?, CellValue::Int(7));
    assert_eq!(
        huge.to_excel_cell(&context)?,
        CellValue::String(huge.to_string())
    );
    Ok(())
}

#[test]
fn floating_point_types_support_numeric_and_string_cells() -> Result<()> {
    let context = context(None);
    for value in [
        f32::from_excel_cell(Some(&CellValue::Int(2)), &context)?,
        f32::from_excel_cell(Some(&CellValue::Float(2.0)), &context)?,
        f32::from_excel_cell(
            Some(&CellValue::Decimal("2.0".parse().expect("valid decimal"))),
            &context,
        )?,
        f32::from_excel_cell(Some(&CellValue::String("2".to_owned())), &context)?,
    ] {
        assert!((value - 2.0).abs() < f32::EPSILON);
    }
    assert!(f32::from_excel_cell(Some(&CellValue::Bool(true)), &context).is_err());
    assert!(f32::from_excel_cell(Some(&CellValue::String("bad".to_owned())), &context).is_err());
    assert!(f32::from_excel_cell(None, &context).is_err());

    for value in [
        f64::from_excel_cell(Some(&CellValue::Int(3)), &context)?,
        f64::from_excel_cell(Some(&CellValue::Float(3.0)), &context)?,
        f64::from_excel_cell(
            Some(&CellValue::Decimal("3.0".parse().expect("valid decimal"))),
            &context,
        )?,
        f64::from_excel_cell(Some(&CellValue::String("3".to_owned())), &context)?,
    ] {
        assert!((value - 3.0).abs() < f64::EPSILON);
    }
    assert!(f64::from_excel_cell(Some(&CellValue::Bool(true)), &context).is_err());
    assert!(f64::from_excel_cell(Some(&CellValue::String("bad".to_owned())), &context).is_err());
    assert!(f64::from_excel_cell(None, &context).is_err());
    assert_eq!(1.25_f32.to_excel_cell(&context)?, CellValue::Float(1.25));
    assert_eq!(2.5_f64.to_excel_cell(&context)?, CellValue::Float(2.5));
    Ok(())
}

#[test]
fn big_decimal_converts_like_java_big_decimal() -> Result<()> {
    let context = context(None);
    let expected: BigDecimal = "123.450".parse().expect("valid decimal");
    assert_eq!(
        BigDecimal::from_excel_cell(Some(&CellValue::Decimal(expected.clone())), &context)?,
        expected
    );
    assert_eq!(
        BigDecimal::from_excel_cell(Some(&CellValue::Int(123)), &context)?,
        BigDecimal::from(123)
    );
    assert_eq!(
        BigDecimal::from_excel_cell(Some(&CellValue::Float(1.25)), &context)?,
        "1.25".parse::<BigDecimal>().expect("valid decimal")
    );
    assert_eq!(
        BigDecimal::from_excel_cell(Some(&CellValue::String("2.50".to_owned())), &context)?,
        "2.50".parse::<BigDecimal>().expect("valid decimal")
    );
    assert!(BigDecimal::from_excel_cell(Some(&CellValue::Float(f64::NAN)), &context).is_err());
    assert!(
        BigDecimal::from_excel_cell(
            Some(&CellValue::String("not-a-number".to_owned())),
            &context
        )
        .is_err()
    );
    assert!(BigDecimal::from_excel_cell(Some(&CellValue::Bool(true)), &context).is_err());
    assert!(BigDecimal::from_excel_cell(None, &context).is_err());
    assert_eq!(
        expected.to_excel_cell(&context)?,
        CellValue::Decimal(expected)
    );
    Ok(())
}

#[test]
fn date_and_datetime_conversion_honors_formats() -> Result<()> {
    let date = NaiveDate::from_ymd_opt(2026, 7, 17).expect("valid date");
    let datetime = date.and_hms_opt(12, 30, 45).expect("valid time");
    let date_context = context(Some("%d/%m/%Y"));
    let datetime_context = context(Some("%d/%m/%Y %H:%M:%S"));

    assert_eq!(
        NaiveDate::from_excel_cell(Some(&CellValue::Date(date)), &date_context)?,
        date
    );
    assert_eq!(
        NaiveDate::from_excel_cell(Some(&CellValue::DateTime(datetime)), &date_context)?,
        date
    );
    assert_eq!(
        NaiveDate::from_excel_cell(
            Some(&CellValue::String("17/07/2026".to_owned())),
            &date_context,
        )?,
        date
    );
    assert!(
        NaiveDate::from_excel_cell(Some(&CellValue::String("bad".to_owned())), &date_context)
            .is_err()
    );
    assert!(NaiveDate::from_excel_cell(Some(&CellValue::Bool(true)), &date_context).is_err());
    assert_eq!(date.to_excel_cell(&date_context)?, CellValue::Date(date));

    assert_eq!(
        NaiveDateTime::from_excel_cell(Some(&CellValue::DateTime(datetime)), &datetime_context)?,
        datetime
    );
    assert_eq!(
        NaiveDateTime::from_excel_cell(Some(&CellValue::Date(date)), &datetime_context)?,
        date.and_hms_opt(0, 0, 0).expect("valid time")
    );
    assert_eq!(
        NaiveDateTime::from_excel_cell(
            Some(&CellValue::String("17/07/2026 12:30:45".to_owned())),
            &datetime_context,
        )?,
        datetime
    );
    assert!(
        NaiveDateTime::from_excel_cell(
            Some(&CellValue::String("bad".to_owned())),
            &datetime_context,
        )
        .is_err()
    );
    assert!(
        NaiveDateTime::from_excel_cell(Some(&CellValue::Bool(true)), &datetime_context).is_err()
    );
    assert_eq!(
        datetime.to_excel_cell(&datetime_context)?,
        CellValue::DateTime(datetime)
    );
    Ok(())
}

#[test]
fn option_conversion_distinguishes_empty_and_present_values() -> Result<()> {
    let context = context(None);
    assert_eq!(Option::<u32>::from_excel_cell(None, &context)?, None);
    assert_eq!(
        Option::<u32>::from_excel_cell(Some(&CellValue::Empty), &context)?,
        None
    );
    assert_eq!(
        Option::<u32>::from_excel_cell(Some(&CellValue::Int(5)), &context)?,
        Some(5)
    );
    assert_eq!(None::<u32>.to_excel_cell(&context)?, CellValue::Empty);
    assert_eq!(Some(5_u32).to_excel_cell(&context)?, CellValue::Int(5));
    Ok(())
}

#[derive(Default)]
struct PrefixConverter;

impl Converter<String> for PrefixConverter {
    fn convert_to_rust_data(&self, context: &ReadConverterContext<'_>) -> Result<String> {
        Ok(format!(
            "{}:{}:{}",
            context.column().field,
            context.convert_context().row_index,
            context.cell().map_or_else(String::new, CellValue::as_text)
        ))
    }

    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> Result<CellValue> {
        Ok(CellValue::String(format!(
            "{}:{}:{}",
            context.column().name,
            context.convert_context().column_index.unwrap_or_default(),
            context.value()
        )))
    }
}

struct UnsupportedConverter;

impl Converter<String> for UnsupportedConverter {}

#[test]
fn custom_converter_contexts_support_both_directions_and_defaults() -> Result<()> {
    let column = ExcelColumn::new("name", "Name", Some(1), 0, Some("text"));
    let context = context(Some("text"));
    let cell = CellValue::String("alice".to_owned());
    let read = ReadConverterContext::new(Some(&cell), &column, &context);
    assert_eq!(read.formula(), None);
    assert_eq!(PrefixConverter.convert_to_rust_data(&read)?, "name:2:alice");
    let formula = FormulaData::new("CONCAT(A1,B1)");
    let formula_read =
        ReadConverterContext::with_formula(Some(&cell), Some(&formula), &column, &context);
    assert_eq!(
        formula_read.formula().map(FormulaData::formula_value),
        Some("CONCAT(A1,B1)")
    );
    let value = "bob".to_owned();
    let write = WriteConverterContext::new(&value, &column, &context);
    assert_eq!(
        PrefixConverter.convert_to_excel_data(&write)?,
        CellValue::String("Name:1:bob".to_owned())
    );

    let empty = ReadConverterContext::new(None, &column, &context);
    assert_eq!(PrefixConverter.convert_to_rust_data(&empty)?, "name:2:");
    assert!(UnsupportedConverter.convert_to_rust_data(&read).is_err());
    assert!(UnsupportedConverter.convert_to_excel_data(&write).is_err());
    assert_eq!(
        UnsupportedConverter.support_excel_type(),
        CellDataType::String
    );
    Ok(())
}

struct BooleanPrefixConverter;

impl Converter<String> for BooleanPrefixConverter {
    fn support_excel_type(&self) -> CellDataType {
        CellDataType::Boolean
    }
}

struct U32Converter;

impl Converter<u32> for U32Converter {}

struct WrongReadTypeConverter;

impl ErasedConverter for WrongReadTypeConverter {
    fn target_type_id(&self) -> std::any::TypeId {
        std::any::TypeId::of::<String>()
    }

    fn target_type_name(&self) -> &'static str {
        "String"
    }

    fn support_excel_type(&self) -> CellDataType {
        CellDataType::String
    }

    fn convert_to_rust_data(&self, _context: &ReadConverterContext<'_>) -> Result<Box<dyn Any>> {
        Ok(Box::new(42_u32))
    }

    fn convert_to_excel_data(
        &self,
        _value: &dyn Any,
        _column: &ExcelColumn,
        _context: &ConvertContext,
    ) -> Result<CellValue> {
        Ok(CellValue::Empty)
    }
}

#[test]
fn converter_registry_matches_java_keys_precedence_and_contract_errors() -> Result<()> {
    let column = ExcelColumn::new("name", "Name", Some(0), 0, None);
    let context = context(None);
    let cell = CellValue::String("alice".to_owned());
    let read = ReadConverterContext::new(Some(&cell), &column, &context);

    let mut registry = ConverterRegistry::default();
    assert!(registry.is_empty());
    assert_eq!(format!("{registry:?}"), "[]");
    assert_eq!(registry.convert_to_rust_data::<String>(&read)?, None);
    assert_eq!(
        registry.convert_to_excel_data(&"bob".to_owned(), &column, &context)?,
        None
    );

    registry.register::<String, _>(PrefixConverter);
    assert!(!registry.is_empty());
    assert!(format!("{registry:?}").contains("String"));
    assert_eq!(registry.clone(), registry);
    assert_eq!(
        registry.convert_to_rust_data::<String>(&read)?,
        Some("name:2:alice".to_owned())
    );
    assert_eq!(
        registry.convert_to_excel_data(&"bob".to_owned(), &column, &context)?,
        Some(CellValue::String("Name:1:bob".to_owned()))
    );

    let mut different_type = ConverterRegistry::default();
    different_type.register::<u32, _>(U32Converter);
    assert_ne!(registry, different_type);
    let mut different_cell_type = ConverterRegistry::default();
    different_cell_type.register::<String, _>(BooleanPrefixConverter);
    assert_ne!(registry, different_cell_type);
    assert_eq!(
        registry
            .merged_with(&different_cell_type)
            .convert_to_rust_data::<String>(&read)?,
        Some("name:2:alice".to_owned())
    );

    let mut unsupported = ConverterRegistry::default();
    unsupported.register::<String, _>(UnsupportedConverter);
    assert!(unsupported.convert_to_rust_data::<String>(&read).is_err());
    assert!(
        unsupported
            .convert_to_excel_data(&"bob".to_owned(), &column, &context)
            .is_err()
    );

    let typed = TypedConverter::<String, PrefixConverter> {
        converter: PrefixConverter,
        marker: std::marker::PhantomData,
    };
    assert!(ErasedConverter::convert_to_excel_data(&typed, &42_u32, &column, &context).is_err());

    let invalid = ConverterRegistry {
        converters: vec![Arc::new(WrongReadTypeConverter)],
    };
    assert!(invalid.convert_to_rust_data::<String>(&read).is_err());
    Ok(())
}

#[test]
fn analysis_context_exposes_sheet_row_and_batch_coordinates() {
    let context = AnalysisContext::new("Users", 3, 9);
    assert_eq!(context.sheet_name(), "Users");
    assert_eq!(context.sheet_no(), 3);
    assert_eq!(context.row_index(), 9);
    assert_eq!(context.batch_index(), 0);
    assert_eq!(context.custom_object(), None);
    assert_eq!(context.custom::<String>(), None);
    assert_eq!(context.with_batch_index(4).batch_index(), 4);

    let custom = CustomReadObject::new("tenant-42".to_owned());
    let shared = custom.clone();
    assert_eq!(custom, shared);
    assert_ne!(custom, CustomReadObject::new("tenant-42".to_owned()));
    assert_eq!(format!("{custom:?}"), "CustomReadObject { .. }");
    assert_eq!(
        custom.downcast_ref::<String>().map(String::as_str),
        Some("tenant-42")
    );
    assert_eq!(custom.downcast_ref::<u32>(), None);

    let context = context.with_custom_object(Some(custom));
    assert!(context.custom_object().is_some());
    assert_eq!(
        context.custom::<String>().map(String::as_str),
        Some("tenant-42")
    );
    assert_eq!(context.custom::<u32>(), None);
    assert_eq!(
        context
            .with_batch_index(7)
            .custom::<String>()
            .map(String::as_str),
        Some("tenant-42")
    );
}

#[derive(Default)]
struct RecordingListener {
    calls: Vec<&'static str>,
}

impl ReadListener<i32> for RecordingListener {
    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        self.calls.push("exception");
        ErrorAction::Continue
    }

    fn invoke_head(
        &mut self,
        _head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        self.calls.push("head");
        Ok(())
    }

    fn invoke(&mut self, _data: i32, _context: &AnalysisContext) -> Result<()> {
        self.calls.push("row");
        Ok(())
    }

    fn extra(&mut self, _extra: &CellExtra, _context: &AnalysisContext) -> Result<()> {
        self.calls.push("extra");
        Ok(())
    }

    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> Result<()> {
        self.calls.push("after");
        Ok(())
    }

    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        self.calls.push("next");
        true
    }
}

struct DefaultListener;

impl ReadListener<i32> for DefaultListener {
    fn invoke(&mut self, _data: i32, _context: &AnalysisContext) -> Result<()> {
        Ok(())
    }
}

#[test]
fn listener_defaults_and_box_forwarding_match_java_lifecycle() -> Result<()> {
    let context = AnalysisContext::new("Users", 0, 1);
    let error = ExcelError::Format("bad".to_owned());
    let extra = CellExtra::new(CellExtraType::Merge, None, 1, 2, 3, 4);
    assert_eq!(extra.extra_type(), CellExtraType::Merge);
    assert_eq!(extra.text(), None);
    assert_eq!(extra.first_row_index(), 1);
    assert_eq!(extra.last_row_index(), 2);
    assert_eq!(extra.first_column_index(), 3);
    assert_eq!(extra.last_column_index(), 4);
    let comment = CellExtra::new(CellExtraType::Comment, Some("note".to_owned()), 0, 0, 0, 0);
    assert_eq!(comment.extra_type(), CellExtraType::Comment);
    assert_eq!(comment.text(), Some("note"));
    assert_eq!(
        CellExtra::new(
            CellExtraType::Hyperlink,
            Some("https://example.com".to_owned()),
            0,
            0,
            0,
            0,
        )
        .extra_type(),
        CellExtraType::Hyperlink
    );
    let mut defaults = DefaultListener;
    assert_eq!(defaults.on_exception(&error, &context), ErrorAction::Stop);
    defaults.invoke_head(&HashMap::new(), &context)?;
    defaults.invoke(1, &context)?;
    defaults.extra(&extra, &context)?;
    defaults.do_after_all_analysed(&context)?;
    assert!(defaults.has_next(&context));

    let mut listener: Box<dyn ReadListener<i32>> = Box::new(RecordingListener::default());
    assert_eq!(
        listener.on_exception(&error, &context),
        ErrorAction::Continue
    );
    listener.invoke_head(&HashMap::new(), &context)?;
    listener.invoke(1, &context)?;
    listener.extra(&extra, &context)?;
    listener.do_after_all_analysed(&context)?;
    assert!(listener.has_next(&context));
    Ok(())
}

#[test]
fn page_listener_flushes_full_partial_and_empty_batches() -> Result<()> {
    let context = AnalysisContext::new("Users", 0, 1);
    let batches = Arc::new(std::sync::Mutex::new(Vec::new()));
    let captured = Arc::clone(&batches);
    let mut listener = PageReadListener::new(0, move |rows: Vec<i32>, context| {
        captured
            .lock()
            .expect("lock")
            .push((rows, context.batch_index()));
        Ok(())
    });
    listener.invoke(1, &context)?;
    listener.do_after_all_analysed(&context)?;
    assert_eq!(&*batches.lock().expect("lock"), &[(vec![1], 0)]);

    let batches = Arc::new(std::sync::Mutex::new(Vec::new()));
    let captured = Arc::clone(&batches);
    let mut listener = PageReadListener::new(2, move |rows: Vec<i32>, context| {
        captured
            .lock()
            .expect("lock")
            .push((rows, context.batch_index()));
        Ok(())
    });
    listener.invoke(1, &context)?;
    listener.invoke(2, &context)?;
    listener.invoke(3, &context)?;
    listener.do_after_all_analysed(&context)?;
    assert_eq!(
        &*batches.lock().expect("lock"),
        &[(vec![1, 2], 0), (vec![3], 1)]
    );
    Ok(())
}

#[test]
fn page_listener_propagates_callback_failures() {
    let context = AnalysisContext::new("Users", 0, 1);
    let mut listener = PageReadListener::new(1, |_rows: Vec<i32>, _context| {
        Err(ExcelError::Format("callback failed".to_owned()))
    });

    let error = listener
        .invoke(1, &context)
        .expect_err("a failed page callback must stop the reader");
    assert_eq!(error.to_string(), "excel format error: callback failed");
}

#[test]
fn every_error_variant_has_actionable_display_text() {
    let data = context(None).invalid(&CellValue::String("bad".to_owned()), "u32");
    assert!(data.to_string().contains("field=value"));
    assert_eq!(
        ExcelError::SheetNotFound("Users".to_owned()).to_string(),
        "worksheet not found: Users"
    );
    assert_eq!(
        ExcelError::Format("bad zip".to_owned()).to_string(),
        "excel format error: bad zip"
    );
    assert_eq!(
        ExcelError::Unsupported("template".to_owned()).to_string(),
        "unsupported operation: template"
    );
    let io_error = ExcelError::from(io::Error::other("disk"));
    assert_eq!(io_error.to_string(), "disk");
    assert_eq!(ErrorAction::SkipRow, ErrorAction::SkipRow);
}

struct DefaultWriteHandler;

impl WriteHandler for DefaultWriteHandler {}

#[test]
fn write_handler_contexts_and_defaults_cover_the_full_lifecycle() -> Result<()> {
    let workbook = WriteWorkbookContext::new("output.xlsx");
    assert_eq!(workbook.path(), std::path::Path::new("output.xlsx"));
    let sheet = WriteSheetContext::new("Users");
    assert_eq!(sheet.sheet_name(), "Users");
    let row = WriteRowContext {
        sheet_name: "Users".to_owned(),
        row_index: 1,
        is_head: false,
    };
    let mut cell = WriteCellContext {
        sheet_name: "Users".to_owned(),
        row_index: 1,
        column_index: 0,
        field: Some("name"),
        is_head: false,
        value: CellValue::String("Alice".to_owned()),
        skip: false,
    };
    let mut handler = DefaultWriteHandler;
    assert_eq!(handler.order(), 0);
    handler.before_workbook(&workbook)?;
    handler.before_sheet(&sheet)?;
    handler.before_row(&row)?;
    handler.before_cell(&mut cell)?;
    handler.after_cell(&cell)?;
    handler.after_row(&row)?;
    handler.after_sheet(&sheet)?;
    handler.after_workbook(&workbook)?;
    Ok(())
}
