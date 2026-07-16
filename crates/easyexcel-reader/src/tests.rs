use std::collections::HashMap;
use std::fs;
use std::io::{BufRead, Read, Write};
use std::sync::Arc;

use base64::Engine;
use calamine::{CellErrorType, ExcelDateTime, ExcelDateTimeType};
use easyexcel_core::{ExcelColumn, IntoExcelCell};
use flate2::read::GzDecoder;
use rust_xlsxwriter::Workbook;
use tempfile::{TempDir, tempdir};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

use super::*;

struct FaultyBufRead;

impl Read for FaultyBufRead {
    fn read(&mut self, _buffer: &mut [u8]) -> std::io::Result<usize> {
        Err(std::io::Error::other("injected probe failure"))
    }
}

impl BufRead for FaultyBufRead {
    fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
        Err(std::io::Error::other("injected probe failure"))
    }

    fn consume(&mut self, _amount: usize) {}
}

fn test_error(error: impl std::fmt::Display) -> ExcelError {
    ExcelError::Format(error.to_string())
}

#[derive(Debug, PartialEq, Eq)]
struct TestRow(String);

impl ExcelRow for TestRow {
    fn schema() -> &'static [ExcelColumn] {
        const COLUMNS: &[ExcelColumn] = &[ExcelColumn::new("value", "Value", Some(0), 0, None)];
        COLUMNS
    }

    fn from_row(row: &RowData) -> Result<Self> {
        let value = row
            .cell(&Self::schema()[0])
            .map_or_else(String::new, CellValue::as_text);
        if value == "conversion-error" {
            return Err(ExcelError::Format("conversion failed".to_owned()));
        }
        Ok(Self(value))
    }

    fn to_row(&self) -> Result<Vec<CellValue>> {
        self.0
            .to_excel_cell(&easyexcel_core::ConvertContext {
                sheet_name: String::new(),
                row_index: 0,
                column_index: Some(0),
                field: "value",
                format: None,
            })
            .map(|value| vec![value])
    }
}

#[derive(Default)]
struct Probe {
    heads: Vec<HashMap<String, usize>>,
    rows: Vec<TestRow>,
    after: Vec<(String, usize, u32)>,
    continue_reading: bool,
    fail_head: bool,
}

impl ReadListener<TestRow> for Probe {
    fn invoke_head(
        &mut self,
        head: &HashMap<String, usize>,
        _context: &AnalysisContext,
    ) -> Result<()> {
        if self.fail_head {
            return Err(ExcelError::Format("head failed".to_owned()));
        }
        self.heads.push(head.clone());
        Ok(())
    }

    fn invoke(&mut self, data: TestRow, _context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        Ok(())
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.after.push((
            context.sheet_name().to_owned(),
            context.sheet_no(),
            context.row_index(),
        ));
        Ok(())
    }

    fn has_next(&mut self, _context: &AnalysisContext) -> bool {
        self.continue_reading
    }
}

struct ErrorProbe {
    action: ErrorAction,
    errors: usize,
}

impl ReadListener<TestRow> for ErrorProbe {
    fn on_exception(&mut self, _error: &ExcelError, _context: &AnalysisContext) -> ErrorAction {
        self.errors += 1;
        self.action
    }

    fn invoke(&mut self, _data: TestRow, _context: &AnalysisContext) -> Result<()> {
        panic!("a conversion failure cannot invoke a row")
    }
}

fn options() -> ReadOptions {
    ReadOptions {
        sheet: SheetSelector::First,
        head_row_number: 1,
        ignore_empty_row: true,
        password: None,
        charset: CsvCharset::default(),
    }
}

fn workbook_fixture() -> Result<(TempDir, std::path::PathBuf)> {
    let directory = tempdir()?;
    let path = directory.path().join("fixture.xlsx");
    let mut workbook = Workbook::new();
    let first = workbook.add_worksheet();
    first.set_name("First").map_err(test_error)?;
    first.write_string(0, 0, "Value").map_err(test_error)?;
    first.write_string(1, 0, "one").map_err(test_error)?;
    let second = workbook.add_worksheet();
    second.set_name("Second").map_err(test_error)?;
    second.write_string(0, 0, "Value").map_err(test_error)?;
    second.write_string(1, 0, "two").map_err(test_error)?;
    workbook.save(&path).map_err(test_error)?;
    Ok((directory, path))
}

fn rewrite_first_sheet(source: &Path, destination: &Path, replacement: &str) -> Result<()> {
    let mut archive = ZipArchive::new(fs::File::open(source)?).map_err(test_error)?;
    let mut writer = ZipWriter::new(fs::File::create(destination)?);
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).map_err(test_error)?;
        let name = entry.name().to_owned();
        let options = SimpleFileOptions::default().compression_method(entry.compression());
        if entry.is_dir() {
            writer.add_directory(name, options).map_err(test_error)?;
            continue;
        }
        writer.start_file(&name, options).map_err(test_error)?;
        if name == "xl/worksheets/sheet1.xml" {
            writer
                .write_all(replacement.as_bytes())
                .map_err(test_error)?;
        } else {
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes)?;
            writer.write_all(&bytes)?;
        }
    }
    writer.finish().map_err(test_error)?;
    Ok(())
}

fn worksheet_xml(cells: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1">{cells}</row></sheetData>
</worksheet>"#
    )
}

fn column_name(index: u32) -> String {
    let mut value = index + 1;
    let mut name = String::new();
    while value > 0 {
        let remainder = ((value - 1) % 26) as u8;
        name.insert(0, char::from(b'A' + remainder));
        value = (value - 1) / 26;
    }
    name
}

fn encode_csv_fixture(encoding: &'static encoding_rs::Encoding, value: &str) -> Vec<u8> {
    if encoding == encoding_rs::UTF_16BE {
        value.encode_utf16().flat_map(u16::to_be_bytes).collect()
    } else if encoding == encoding_rs::UTF_16LE {
        value.encode_utf16().flat_map(u16::to_le_bytes).collect()
    } else {
        let (encoded, actual, had_errors) = encoding.encode(value);
        assert_eq!(actual, encoding);
        assert!(!had_errors);
        encoded.into_owned()
    }
}

#[test]
fn calamine_values_map_to_every_core_cell_variant() {
    let datetime = ExcelDateTime::new(46_120.5, ExcelDateTimeType::DateTime, false);
    let invalid_datetime = ExcelDateTime::new(f64::MAX, ExcelDateTimeType::DateTime, false);
    let duration = ExcelDateTime::new(1.5, ExcelDateTimeType::TimeDelta, false);
    let cases = [
        (DataRef::Empty, CellValue::Empty),
        (
            DataRef::String("owned".to_owned()),
            CellValue::String("owned".to_owned()),
        ),
        (
            DataRef::SharedString("shared"),
            CellValue::String("shared".to_owned()),
        ),
        (
            DataRef::DateTimeIso("2026-01-01".to_owned()),
            CellValue::String("2026-01-01".to_owned()),
        ),
        (
            DataRef::DurationIso("PT1H".to_owned()),
            CellValue::String("PT1H".to_owned()),
        ),
        (DataRef::Bool(true), CellValue::Bool(true)),
        (DataRef::Int(7), CellValue::Int(7)),
        (DataRef::Float(1.25), CellValue::Float(1.25)),
        (DataRef::DateTime(duration), CellValue::Float(1.5)),
        (
            DataRef::DateTime(invalid_datetime),
            CellValue::Float(f64::MAX),
        ),
        (
            DataRef::Error(CellErrorType::Div0),
            CellValue::Error("Div0".to_owned()),
        ),
    ];
    for (input, expected) in cases {
        assert_eq!(from_calamine(&input), expected);
        assert_eq!(from_data(&Data::from(input)), expected);
    }
    assert!(matches!(
        from_calamine(&DataRef::DateTime(datetime)),
        CellValue::DateTime(_)
    ));
    assert!(matches!(
        from_data(&Data::DateTime(datetime)),
        CellValue::DateTime(_)
    ));
}

#[test]
fn helpers_preserve_diagnostics_and_xlsx_column_limits() {
    assert_eq!(ReadOptions::default(), options());
    assert_eq!(SheetSelector::default(), SheetSelector::First);
    assert_eq!(to_column_index(0).expect("column"), 0);
    assert_eq!(
        to_column_index(u32::from(u16::MAX)).expect("column"),
        usize::from(u16::MAX)
    );
    assert!(to_column_index(u32::from(u16::MAX) + 1).is_err());
    assert_eq!(
        format_error("broken").to_string(),
        "excel format error: broken"
    );
    assert!(!is_compound_document(&mut FaultyBufRead));
}

#[test]
fn legacy_range_read_preserves_coordinates_headers_and_empty_sheets() -> Result<()> {
    let mut range = Range::new((2, 1), (3, 1));
    range.set_value((2, 1), Data::String("Value".to_owned()));
    range.set_value((3, 1), Data::String("one".to_owned()));
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_range(
        &range,
        2,
        "Legacy",
        &ReadOptions {
            head_row_number: 3,
            ..options()
        },
        &mut TypedRowConsumer::<TestRow> {
            listener: &mut probe,
        },
    )?;
    assert_eq!(probe.heads[0].get("Value"), Some(&1));
    assert_eq!(probe.rows, vec![TestRow(String::new())]);
    assert_eq!(probe.after, vec![("Legacy".to_owned(), 2, 3)]);

    read_range(
        &Range::empty(),
        3,
        "Empty",
        &options(),
        &mut TypedRowConsumer::<TestRow> {
            listener: &mut probe,
        },
    )?;
    assert_eq!(probe.after.last(), Some(&("Empty".to_owned(), 3, 0)));

    let invalid_column = Range::new((0, u32::from(u16::MAX) + 1), (0, u32::from(u16::MAX) + 1));
    assert!(
        read_range(
            &invalid_column,
            0,
            "Invalid",
            &options(),
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut probe,
            },
        )
        .is_err()
    );

    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_range(
            &range,
            0,
            "Legacy",
            &ReadOptions {
                head_row_number: 3,
                ..options()
            },
            &mut TypedRowConsumer::<TestRow> {
                listener: &mut failing_head,
            },
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn reads_java_easyexcel_legacy_multisheet_fixture() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("java-multiplesheets.xls");
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(include_str!("fixtures/java-multiplesheets.xls.gz.b64").trim())
        .map_err(test_error)?;
    let mut decoder = GzDecoder::new(compressed.as_slice());
    let mut workbook = Vec::new();
    decoder.read_to_end(&mut workbook)?;
    fs::write(&path, workbook)?;
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xls::<TestRow, _>(
        &path,
        &ReadOptions {
            sheet: SheetSelector::All,
            ..options()
        },
        &mut probe,
    )?;
    assert_eq!(
        probe.rows,
        (1..=6)
            .map(|index| TestRow(format!("表{index}数据")))
            .collect::<Vec<_>>()
    );
    assert_eq!(
        probe.after,
        (0..6)
            .map(|index| (format!("Sheet{}", index + 1), index, 1))
            .collect::<Vec<_>>()
    );
    assert_eq!(probe.heads.len(), 6);
    for (index, head) in probe.heads.iter().enumerate() {
        assert_eq!(head.get(&format!("表{}头", index + 1)), Some(&0));
    }
    assert!(
        read_xls::<TestRow, _>(
            &path,
            &ReadOptions {
                sheet: SheetSelector::Index(99),
                ..options()
            },
            &mut probe,
        )
        .is_err()
    );
    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(read_xls::<TestRow, _>(&path, &options(), &mut failing_head).is_err());
    let invalid = directory.path().join("invalid.xls");
    fs::write(&invalid, b"not an XLS workbook")?;
    assert!(read_xls::<TestRow, _>(&invalid, &options(), &mut probe).is_err());
    Ok(())
}

#[test]
fn reads_java_easyexcel_encrypted_xlsx_fixture() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("java-encrypt07.xlsx");
    let compressed = base64::engine::general_purpose::STANDARD
        .decode(include_str!("fixtures/java-encrypt07.xlsx.gz.b64").trim())
        .map_err(test_error)?;
    let mut decoder = GzDecoder::new(compressed.as_slice());
    let mut workbook = Vec::new();
    decoder.read_to_end(&mut workbook)?;
    assert!(is_compound_document(&mut workbook.as_slice()));
    assert!(!is_compound_document(&mut &workbook[..4]));
    assert!(!is_compound_document(&mut &b"not-cfb!"[..]));
    fs::write(&path, workbook)?;

    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            password: Some("123456".to_owned()),
            ..options()
        },
        &mut probe,
    )?;
    assert!(read_xlsx::<TestRow, _>(&path, &options(), &mut probe).is_err());
    assert_eq!(
        probe.rows,
        (0..10)
            .map(|index| TestRow(format!("姓名{index}")))
            .collect::<Vec<_>>()
    );
    assert_eq!(probe.heads[0].get("姓名"), Some(&0));
    assert_eq!(probe.after, vec![("0".to_owned(), 0, 10)]);
    Ok(())
}

#[test]
fn sheet_selection_supports_first_index_name_all_and_missing_values() -> Result<()> {
    let (_directory, path) = workbook_fixture()?;
    let workbook: Xlsx<_> = open_workbook(path).map_err(test_error)?;
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::First)?,
        vec![(0, "First".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::Index(1))?,
        vec![(1, "Second".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::Name("Second".to_owned()))?,
        vec![(1, "Second".to_owned())]
    );
    assert_eq!(
        selected_sheet_names(&workbook, &SheetSelector::All)?.len(),
        2
    );
    assert!(selected_sheet_names(&workbook, &SheetSelector::Index(2)).is_err());
    assert!(selected_sheet_names(&workbook, &SheetSelector::Name("Missing".to_owned())).is_err());
    assert!(select_sheet_names(Vec::new(), &SheetSelector::First).is_err());

    let legacy = || {
        vec![
            ("First".to_owned(), Range::empty()),
            ("Second".to_owned(), Range::empty()),
        ]
    };
    let first = select_xls_sheets(legacy(), &SheetSelector::First)?;
    assert_eq!((first[0].0, first[0].1.as_str()), (0, "First"));
    let second = select_xls_sheets(legacy(), &SheetSelector::Index(1))?;
    assert_eq!((second[0].0, second[0].1.as_str()), (1, "Second"));
    let named = select_xls_sheets(legacy(), &SheetSelector::Name("Second".to_owned()))?;
    assert_eq!((named[0].0, named[0].1.as_str()), (1, "Second"));
    assert_eq!(select_xls_sheets(legacy(), &SheetSelector::All)?.len(), 2);
    assert!(select_xls_sheets(legacy(), &SheetSelector::Index(2)).is_err());
    assert!(select_xls_sheets(legacy(), &SheetSelector::Name("Missing".to_owned())).is_err());
    assert!(select_xls_sheets(Vec::new(), &SheetSelector::First).is_err());
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn csv_read_uses_typed_lifecycle_single_sheet_selection_and_flexible_rows() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("fixture.csv");
    fs::write(&path, b"\xEF\xBB\xBFValue,Extra\r\none,1\r\ntwo\r\n")?;
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_csv::<TestRow, _>(&path, &options(), &mut probe)?;
    assert_eq!(
        probe.rows,
        vec![TestRow("one".to_owned()), TestRow("two".to_owned())]
    );
    assert_eq!(probe.heads[0].get("Value"), Some(&0));
    assert_eq!(probe.after, vec![("Sheet1".to_owned(), 0, 2)]);

    assert_eq!(csv_sheet_name(&SheetSelector::First)?, "Sheet1");
    assert_eq!(csv_sheet_name(&SheetSelector::Index(0))?, "Sheet1");
    assert_eq!(csv_sheet_name(&SheetSelector::All)?, "Sheet1");
    assert_eq!(
        csv_sheet_name(&SheetSelector::Name("Custom".to_owned()))?,
        "Custom"
    );
    assert!(csv_sheet_name(&SheetSelector::Index(1)).is_err());
    assert_eq!(csv_row_index(0)?, 0);
    if usize::BITS > 32 {
        assert!(csv_row_index(usize::try_from(u64::from(u32::MAX) + 1).unwrap()).is_err());
    }

    let malformed_utf8 = directory.path().join("malformed-utf8.csv");
    fs::write(&malformed_utf8, [0xff])?;
    let mut replacement_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_csv::<TestRow, _>(
        &malformed_utf8,
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut replacement_probe,
    )?;
    assert_eq!(replacement_probe.rows, vec![TestRow("�".to_owned())]);
    assert!(
        read_csv::<TestRow, _>(
            &path,
            &ReadOptions {
                sheet: SheetSelector::Index(1),
                ..options()
            },
            &mut probe
        )
        .is_err()
    );
    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(read_csv::<TestRow, _>(&path, &options(), &mut failing_head).is_err());
    let record = csv::StringRecord::from(vec!["value"]);
    let mut record_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_csv_records::<TestRow, _>(
        &mut [Ok(record.clone())].into_iter(),
        0,
        "Sheet1",
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut record_probe,
    )?;
    assert_eq!(record_probe.rows, vec![TestRow("value".to_owned())]);
    read_csv_records::<TestRow, _>(
        &mut [Ok(record.clone()), Ok(record.clone())].into_iter(),
        0,
        "Sheet1",
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut record_probe,
    )?;
    assert_eq!(record_probe.rows.len(), 3);
    let mut invalid_utf8_reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_reader([0xFF].as_slice());
    assert!(
        read_csv_records::<TestRow, _>(
            &mut invalid_utf8_reader.records(),
            0,
            "Sheet1",
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut record_probe,
        )
        .is_err()
    );
    if usize::BITS > 32 {
        assert!(
            read_csv_records::<TestRow, _>(
                &mut [Ok(record.clone())].into_iter(),
                usize::try_from(u64::from(u32::MAX) + 1).unwrap(),
                "Sheet1",
                &ReadOptions {
                    head_row_number: 0,
                    ..options()
                },
                &mut probe
            )
            .is_err()
        );
    }
    assert!(
        read_csv_records::<TestRow, _>(
            &mut [Ok(record.clone()), Ok(record)].into_iter(),
            usize::MAX,
            "Sheet1",
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut probe
        )
        .is_err()
    );
    assert!(
        read_csv::<TestRow, _>(
            &directory.path().join("missing.csv"),
            &options(),
            &mut probe
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn csv_read_decodes_java_charset_names_and_strips_matching_bom() -> Result<()> {
    let directory = tempdir()?;
    for (name, encoding, bom) in [
        ("utf-8", encoding_rs::UTF_8, b"\xEF\xBB\xBF".as_slice()),
        ("GBK", encoding_rs::GBK, b"".as_slice()),
        ("UTF-16BE", encoding_rs::UTF_16BE, b"\xFE\xFF".as_slice()),
        ("UTF-16LE", encoding_rs::UTF_16LE, b"\xFF\xFE".as_slice()),
    ] {
        let path = directory
            .path()
            .join(format!("{}.csv", name.to_lowercase()));
        let mut bytes = bom.to_vec();
        bytes.extend_from_slice(&encode_csv_fixture(encoding, "Value\r\n姓名\r\n"));
        fs::write(&path, bytes)?;

        let mut probe = Probe {
            continue_reading: true,
            ..Probe::default()
        };
        read_csv::<TestRow, _>(
            &path,
            &ReadOptions {
                charset: CsvCharset::new(name),
                ..options()
            },
            &mut probe,
        )?;
        assert_eq!(
            probe.rows,
            vec![TestRow("姓名".to_owned())],
            "charset {name}"
        );
    }

    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    let error = read_csv::<TestRow, _>(
        &directory.path().join("utf-8.csv"),
        &ReadOptions {
            charset: CsvCharset::new("not-a-charset"),
            ..options()
        },
        &mut probe,
    )
    .expect_err("unknown charset must be rejected");
    assert!(matches!(error, ExcelError::Unsupported(_)));
    Ok(())
}

#[test]
fn reads_java_easyexcel_official_csv_bom_fixtures() -> Result<()> {
    let directory = tempdir()?;
    for (name, fixture) in [
        (
            "no-bom.csv",
            include_str!("fixtures/java-bom-no-bom.csv.b64"),
        ),
        (
            "office-bom.csv",
            include_str!("fixtures/java-bom-office-bom.csv.b64"),
        ),
    ] {
        let path = directory.path().join(name);
        fs::write(
            &path,
            base64::engine::general_purpose::STANDARD
                .decode(fixture.trim())
                .map_err(test_error)?,
        )?;
        let mut probe = Probe {
            continue_reading: true,
            ..Probe::default()
        };
        read_csv::<TestRow, _>(&path, &options(), &mut probe)?;
        assert_eq!(probe.rows.len(), 10);
        assert_eq!(probe.rows[0], TestRow("姓名0".to_owned()));
        assert_eq!(probe.rows[9], TestRow("姓名9".to_owned()));
    }
    Ok(())
}

#[test]
fn row_processing_handles_headers_skips_data_and_listener_failures() -> Result<()> {
    let mut headers = Arc::new(HashMap::new());
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    process_row::<TestRow>(
        0,
        "First",
        0,
        vec![CellValue::String("Value".to_owned()), CellValue::Empty],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.heads[0].get("Value"), Some(&0));

    process_row::<TestRow>(
        0,
        "First",
        1,
        vec![CellValue::String("one".to_owned())],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows, vec![TestRow("one".to_owned())]);

    process_row::<TestRow>(
        0,
        "First",
        2,
        vec![CellValue::Empty],
        &options(),
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows.len(), 1);

    let two_header_rows = ReadOptions {
        head_row_number: 2,
        ..options()
    };
    process_row::<TestRow>(
        0,
        "First",
        0,
        vec![CellValue::String("ignored".to_owned())],
        &two_header_rows,
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows.len(), 1);

    probe.continue_reading = false;
    let include_empty = ReadOptions {
        ignore_empty_row: false,
        ..options()
    };
    process_row::<TestRow>(
        0,
        "First",
        3,
        vec![CellValue::Empty],
        &include_empty,
        &mut headers,
        &mut probe,
    )?;
    assert_eq!(probe.rows.len(), 1);

    let mut failing_head = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("Value".to_owned())],
            &options(),
            &mut headers,
            &mut failing_head
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn conversion_error_actions_continue_skip_or_stop() -> Result<()> {
    let mut headers = Arc::new(HashMap::new());
    let read_options = ReadOptions {
        head_row_number: 0,
        ignore_empty_row: false,
        ..options()
    };
    for action in [ErrorAction::Continue, ErrorAction::SkipRow] {
        let mut listener = ErrorProbe { action, errors: 0 };
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("conversion-error".to_owned())],
            &read_options,
            &mut headers,
            &mut listener,
        )?;
        assert_eq!(listener.errors, 1);
    }
    let mut listener = ErrorProbe {
        action: ErrorAction::Stop,
        errors: 0,
    };
    assert!(
        process_row::<TestRow>(
            0,
            "First",
            0,
            vec![CellValue::String("conversion-error".to_owned())],
            &read_options,
            &mut headers,
            &mut listener
        )
        .is_err()
    );
    assert_eq!(listener.errors, 1);
    Ok(())
}

#[test]
#[allow(clippy::too_many_lines)]
fn public_reader_streams_all_sheets_and_reports_invalid_workbooks() -> Result<()> {
    let (fixture_directory, path) = workbook_fixture()?;
    let mut probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &path,
        &ReadOptions {
            sheet: SheetSelector::All,
            ..options()
        },
        &mut probe,
    )?;
    assert_eq!(
        probe.rows,
        vec![TestRow("one".to_owned()), TestRow("two".to_owned())]
    );
    assert_eq!(
        probe.after,
        vec![("First".to_owned(), 0, 1), ("Second".to_owned(), 1, 1)]
    );

    let mut missing = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &path,
            &ReadOptions {
                sheet: SheetSelector::Index(99),
                ..options()
            },
            &mut missing,
        )
        .is_err()
    );

    let mut failing_transition = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(&path, &options(), &mut failing_transition).is_err(),
        "a header error emitted while advancing rows must propagate"
    );

    let single_path = fixture_directory.path().join("single.xlsx");
    let mut workbook = Workbook::new();
    workbook
        .add_worksheet()
        .write_string(0, 0, "Value")
        .map_err(test_error)?;
    workbook.save(&single_path).map_err(test_error)?;
    let mut failing_final = Probe {
        continue_reading: true,
        fail_head: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(&single_path, &options(), &mut failing_final).is_err(),
        "a header error emitted at end-of-sheet must propagate"
    );

    let mut opened: Xlsx<_> = open_workbook(&path).map_err(test_error)?;
    let mut direct = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    let mut consumer = TypedRowConsumer::<TestRow> {
        listener: &mut direct,
    };
    assert!(read_sheet(&mut opened, 0, "Missing", &options(), &mut consumer).is_err());
    let mut public_opened = open_xlsx(&path, None)?;
    assert!(read_sheet(&mut public_opened, 0, "Missing", &options(), &mut consumer,).is_err());

    let empty_path = fixture_directory.path().join("empty.xlsx");
    let mut empty_workbook = Workbook::new();
    empty_workbook.add_worksheet();
    empty_workbook.save(&empty_path).map_err(test_error)?;
    let mut empty_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(&empty_path, &options(), &mut empty_probe)?;
    assert!(empty_probe.rows.is_empty());
    assert_eq!(empty_probe.after, vec![("Sheet1".to_owned(), 0, 0)]);

    let out_of_order_path = fixture_directory.path().join("out-of-order.xlsx");
    let out_of_order_xml = worksheet_xml(
        r#"<c r="B1" t="inlineStr"><is><t>second</t></is></c>
<c r="A1" t="inlineStr"><is><t>first</t></is></c>"#,
    );
    rewrite_first_sheet(&path, &out_of_order_path, &out_of_order_xml)?;
    let mut out_of_order_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    read_xlsx::<TestRow, _>(
        &out_of_order_path,
        &ReadOptions {
            head_row_number: 0,
            ..options()
        },
        &mut out_of_order_probe,
    )?;
    assert_eq!(out_of_order_probe.rows, vec![TestRow("first".to_owned())]);

    let wide_path = fixture_directory.path().join("wide.xlsx");
    let wide_column = column_name(u32::from(u16::MAX) + 1);
    let wide_xml = worksheet_xml(&format!(
        r#"<c r="{wide_column}1" t="inlineStr"><is><t>wide</t></is></c>"#
    ));
    rewrite_first_sheet(&path, &wide_path, &wide_xml)?;
    let mut wide_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &wide_path,
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut wide_probe,
        )
        .is_err()
    );

    let truncated_path = fixture_directory.path().join("truncated.xlsx");
    rewrite_first_sheet(
        &path,
        &truncated_path,
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>first</t></is></c>"#,
    )?;
    let mut truncated_probe = Probe {
        continue_reading: true,
        ..Probe::default()
    };
    assert!(
        read_xlsx::<TestRow, _>(
            &truncated_path,
            &ReadOptions {
                head_row_number: 0,
                ..options()
            },
            &mut truncated_probe,
        )
        .is_err()
    );

    let directory = tempdir()?;
    let invalid = directory.path().join("invalid.xlsx");
    fs::write(&invalid, b"not an xlsx")?;
    assert!(read_xlsx::<TestRow, _>(&invalid, &options(), &mut probe).is_err());
    assert!(
        read_xlsx::<TestRow, _>(
            &directory.path().join("missing.xlsx"),
            &options(),
            &mut probe,
        )
        .is_err()
    );
    assert!(read_xlsx::<TestRow, _>(directory.path(), &options(), &mut probe).is_err());
    let invalid_encrypted = directory.path().join("invalid-encrypted.xlsx");
    fs::write(
        &invalid_encrypted,
        [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1],
    )?;
    assert!(
        read_xlsx::<TestRow, _>(
            &invalid_encrypted,
            &ReadOptions {
                password: Some("123456".to_owned()),
                ..options()
            },
            &mut probe,
        )
        .is_err()
    );
    Ok(())
}
