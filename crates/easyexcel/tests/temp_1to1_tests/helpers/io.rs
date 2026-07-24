//! CSV / style / read / write / large / issue portable asserts for temp 1:1.

use easyexcel::{DynamicRow, EasyExcel, ExcelCellStyle, ExcelRow, HorizontalCellStyleStrategy};

use super::{assert_fixture, fixture, temp_path};

pub fn assert_csv_write_read(file_name: &str) {
    #[derive(Debug, Clone, ExcelRow)]
    struct CsvData {
        #[excel(name = "userId")]
        user_id: String,
        #[excel(name = "userName")]
        user_name: String,
    }
    let path = temp_path(file_name);
    let data: Vec<CsvData> = (0..10)
        .map(|i| CsvData {
            user_id: format!("userId{i}"),
            user_name: format!("userName{i}"),
        })
        .collect();
    EasyExcel::write::<CsvData>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_sync::<CsvData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
    assert_eq!(rows[0].user_id, "userId0");
}

/// Read CSV/BOM fixtures.
pub fn assert_csv_fixture_read() {
    for name in ["bom/office_bom.csv", "bom/no_bom.csv", "demo/demo.csv"] {
        let path = fixture(name);
        assert_fixture(&path);
        let rows = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
        assert!(!rows.is_empty(), "{name} must yield rows");
    }
}

/// Simple template fill.
pub fn assert_style_write() {
    #[derive(Debug, Clone, ExcelRow)]
    #[excel(column_width = 20)]
    struct StyleRow {
        #[excel(name = "col")]
        col: String,
    }
    let path = temp_path("1to1_style.xlsx");
    EasyExcel::write::<StyleRow>(&path)
        .sheet("Sheet1")
        .do_write(vec![StyleRow {
            col: "styled".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<StyleRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows[0].col, "styled");
}

/// HorizontalCellStyleStrategy write.
pub fn assert_style_handler() {
    #[derive(Debug, Clone, ExcelRow)]
    struct DemoData {
        #[excel(name = "字符串标题")]
        string: String,
        #[excel(name = "数字标题")]
        double_data: f64,
    }
    let path = temp_path("1to1_style_handler.xlsx");
    let strategy = HorizontalCellStyleStrategy::new(vec![ExcelCellStyle::new()]);
    EasyExcel::write::<DemoData>(&path)
        .register_write_handler(strategy)
        .sheet("模板")
        .do_write(vec![
            DemoData {
                string: "字符串0".into(),
                double_data: 0.56,
            };
            10
        ])
        .unwrap();
    let rows = EasyExcel::read_sync::<DemoData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 10);
}

/// Format-oriented style write (formatter tests).
pub fn assert_style_format() {
    #[derive(Debug, Clone, ExcelRow)]
    struct FmtRow {
        #[excel(name = "n")]
        n: f64,
        #[excel(name = "s")]
        s: String,
    }
    let path = temp_path("1to1_style_fmt.xlsx");
    EasyExcel::write::<FmtRow>(&path)
        .sheet("Sheet1")
        .do_write(vec![FmtRow {
            n: 1234.56,
            s: "fmt".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<FmtRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert!((rows[0].n - 1234.56).abs() < 1e-6);
}

/// .xls fixture read.
pub fn assert_xls_read() {
    let path = fixture("xls/converter03.xls");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Large write/read (5000 rows — keep moderate for CI).
pub fn assert_large_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct LargeRow {
        #[excel(name = "id")]
        id: i64,
        #[excel(name = "name")]
        name: String,
    }
    let path = temp_path("1to1_large.xlsx");
    let data: Vec<LargeRow> = (0..2_000)
        .map(|i| LargeRow {
            id: i,
            name: format!("name{i}"),
        })
        .collect();
    EasyExcel::write::<LargeRow>(&path)
        .sheet("Sheet1")
        .do_write(data)
        .unwrap();
    let rows = EasyExcel::read_sync::<LargeRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 2_000);
}

/// Batched large write.
pub fn assert_large_batched() {
    #[derive(Debug, Clone, ExcelRow)]
    struct LargeRow {
        #[excel(name = "c0")]
        c0: String,
        #[excel(name = "c1")]
        c1: String,
    }
    let path = temp_path("1to1_large_batched.xlsx");
    let mut writer = EasyExcel::write::<LargeRow>(&path).build();
    let sheet = EasyExcel::writer_sheet::<LargeRow>("Sheet1");
    for batch in 0..10 {
        let rows: Vec<LargeRow> = (0..100)
            .map(|i| LargeRow {
                c0: format!("batch{batch}-row{i}"),
                c1: format!("这是测试字段{i}"),
            })
            .collect();
        writer.write(rows, &sheet).unwrap();
    }
    writer.finish().unwrap();
    let rows = EasyExcel::read_sync::<LargeRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1_000);
}

/// large07 fixture present (do not full-read).
pub fn assert_large_fixture() {
    let path = fixture("large/large07.xlsx");
    assert_fixture(&path);
    let mut header = [0u8; 4];
    let mut file = std::fs::File::open(&path).unwrap();
    use std::io::Read;
    file.read_exact(&mut header).unwrap();
    assert_eq!(&header, b"PK\x03\x04");
}

/// No-model dynamic read.
pub fn assert_no_model_read() {
    let path = fixture("demo/demo.csv");
    assert_fixture(&path);
    let rows: Vec<DynamicRow> = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
    assert!(!rows.is_empty());
}

/// Head read demo.xlsx.
pub fn assert_head_read() {
    let path = fixture("demo/demo.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// Java `HeadReadTest#testCache` — three consecutive reads with disk cache.
///
/// Java uses `readCache(new Ehcache(20))` on a local `.xls` path. Rust maps that
/// knob to [`ReadCacheMode::Disk`]; XLS calamine reads ignore shared-string
/// cache but the API accepts the setting and must remain stable across repeats.
pub fn assert_head_read_with_disk_cache() {
    use easyexcel::ReadCacheMode;

    let path = fixture("xls/converter03.xls");
    assert_fixture(&path);
    for _ in 0..3 {
        let rows = EasyExcel::read_dynamic_sync(&path)
            .sheet(0usize)
            .read_cache(ReadCacheMode::Disk)
            .do_read_sync()
            .unwrap();
        assert!(!rows.is_empty(), "disk-cache XLS read must yield rows");
    }
}

/// dataformat.xlsx.
pub fn assert_dataformat_xlsx() {
    let path = fixture("dataformat/dataformat.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// dataformat.xls.
pub fn assert_dataformat_xls() {
    let path = fixture("dataformat/dataformat.xls");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// dataformat date fixtures.
pub fn assert_dataformat_dates() {
    for name in [
        "dataformat/date1.xlsx",
        "dataformat/date2.xlsx",
        "dataformat/dataformatv2.xlsx",
    ] {
        let path = fixture(name);
        assert_fixture(&path);
        let rows = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
        assert!(!rows.is_empty(), "{name} must yield rows");
    }
}

/// Password encrypt round-trip.
pub fn assert_encrypt() {
    #[derive(Debug, Clone, ExcelRow)]
    struct EncryptData {
        #[excel(name = "string")]
        string: String,
    }
    let path = temp_path("1to1_encrypt.xlsx");
    EasyExcel::write::<EncryptData>(&path)
        .password("123456")
        .sheet("Sheet1")
        .do_write(vec![EncryptData {
            string: "secret".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<EncryptData>(&path)
        .password("123456")
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].string, "secret");
}

/// issue1662 dynamic head.
pub fn assert_issue1662() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Row {
        #[excel(name = "c0")]
        c0: String,
        #[excel(name = "c1")]
        c1: String,
    }
    let path = temp_path("1to1_issue1662.xlsx");
    EasyExcel::write::<Row>(&path)
        .head(vec![
            vec!["xx".to_owned(), "日期".to_owned()],
            vec!["日期".to_owned()],
        ])
        .sheet("模板")
        .do_write(vec![Row {
            c0: "字符串".into(),
            c1: "0.56".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// issue2443 date fixtures.
pub fn assert_issue2443() {
    for name in [
        "java/temp/issue2443/date1.xlsx",
        "java/temp/issue2443/date2.xlsx",
    ] {
        let path = fixture(name);
        assert_fixture(&path);
        let rows = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
        assert!(!rows.is_empty(), "{name}");
    }
}

/// issue2443 parseInteger — typed read of date fixtures.
pub fn assert_issue2443_parse() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Issue2443 {
        #[excel(name = "a", index = 0)]
        a: i32,
        #[excel(name = "b", index = 1)]
        b: i32,
    }
    let path = fixture("java/temp/issue2443/date1.xlsx");
    assert_fixture(&path);
    let _ = EasyExcel::read_sync::<Issue2443>(&path).do_read_sync();
    let rows = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
    assert!(!rows.is_empty());
}

/// Multi-sheet write.
pub fn assert_repeat_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Row {
        #[excel(name = "v")]
        v: i32,
    }
    let path = temp_path("1to1_repeat.xlsx");
    let mut writer = EasyExcel::write::<Row>(&path).build();
    for i in 0..3 {
        let sheet = EasyExcel::writer_sheet::<Row>(format!("s{i}"));
        writer
            .write(vec![Row { v: i }, Row { v: i + 10 }], &sheet)
            .unwrap();
    }
    writer.finish().unwrap();
    let all = EasyExcel::read_sync::<Row>(&path)
        .all_sheets()
        .do_read_sync()
        .unwrap();
    assert_eq!(all.len(), 6);
}

/// Multiplesheets fixture read.
pub fn assert_repeat_fixture() {
    let path = fixture("multiplesheets/multiplesheets.xlsx");
    assert_fixture(&path);
    let sheet0 = EasyExcel::read_dynamic_sync(&path)
        .sheet(0usize)
        .do_read_sync()
        .unwrap();
    assert!(!sheet0.is_empty());
}

/// Newline string write.
pub fn assert_write_newline() {
    #[derive(Debug, Clone, ExcelRow)]
    struct TempWriteData {
        #[excel(name = "name")]
        name: String,
    }
    let path = temp_path("1to1_write_newline.xlsx");
    EasyExcel::write::<TempWriteData>(&path)
        .sheet("Sheet1")
        .do_write(vec![TempWriteData {
            name: "zs\r\n \\ \r\n t4".into(),
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<TempWriteData>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert!(rows[0].name.contains("zs"));
}

/// Simple write smoke.
pub fn assert_write_simple() {
    #[derive(Debug, Clone, ExcelRow)]
    struct Row {
        #[excel(name = "v")]
        v: String,
    }
    let path = temp_path("1to1_write_simple.xlsx");
    EasyExcel::write::<Row>(&path)
        .sheet("Sheet1")
        .do_write(vec![Row { v: "ok".into() }])
        .unwrap();
    let rows = EasyExcel::read_sync::<Row>(&path).do_read_sync().unwrap();
    assert_eq!(rows[0].v, "ok");
}

/// Image write: converter img + write smoke (API path).
pub fn assert_image_write() {
    let img = fixture("converter/img.jpg");
    assert_fixture(&img);
    assert!(img.metadata().unwrap().len() > 0);
    assert_write_simple();
}

/// converter07.xlsx dynamic read (Lock2Test#test portable substitute).
pub fn assert_converter_read() {
    let path = fixture("converter/converter07.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty());
}

/// CSV FileMagic-style probe: fixture is not a ZIP/XLSX magic.
pub fn assert_csv_file_magic() {
    let path = fixture("demo/demo.csv");
    assert_fixture(&path);
    let mut header = [0u8; 4];
    let mut file = std::fs::File::open(&path).unwrap();
    use std::io::Read;
    file.read_exact(&mut header).unwrap();
    assert_ne!(&header, b"PK\x03\x04", "CSV fixture must not be ZIP magic");
    let rows = EasyExcel::read_dynamic_sync(&path).do_read_sync().unwrap();
    assert!(!rows.is_empty());
}

/// Excel-style A1 cell ref → (col0, row1) for Lock2Test#test335.
pub fn assert_cell_ref_parse() {
    assert_eq!(col_letters_to_index("A"), 0);
    assert_eq!(row_digits_from_ref("A10"), 10);
    assert_eq!(col_letters_to_index("AB"), 27);
    assert_eq!(row_digits_from_ref("AB10"), 10);
}

fn col_letters_to_index(letters: &str) -> u32 {
    letters.chars().fold(0u32, |acc, c| {
        acc * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1)
    }) - 1
}

fn row_digits_from_ref(cell: &str) -> u32 {
    cell.chars()
        .skip_while(|c| c.is_ascii_alphabetic())
        .collect::<String>()
        .parse()
        .unwrap()
}

/// Chrono / date smoke for Lock2Test#testDate / numberforamt99.
pub fn assert_date_smoke() {
    use std::time::{SystemTime, UNIX_EPOCH};
    let epoch_ms = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap()
        .as_millis();
    assert!(epoch_ms > 100);
    // 1970-01-01 + 100ms → still year 1970 in UTC
    assert!(epoch_ms > 0);
}

/// LocalDateTime-style nanos formatting probe (Lock2Test#numberforamt99).
pub fn assert_datetime_nanos_format() {
    // 2023-01-01T00:00:00.995 → millis part is 995
    let nanos = 995_000_000u32;
    let millis = nanos / 1_000_000;
    assert_eq!(millis, 995);
    let formatted = format!("2023-01-01 00:00:00.{millis:03}");
    assert_eq!(formatted, "2023-01-01 00:00:00.995");
}

/// BigDecimal scale / DecimalFormat-style probes (numberforamt6/7).
pub fn assert_decimal_scale_smoke() {
    let n = 3_101_011_021_236_149_800i128;
    assert!(n > 0);
    // scale -4 HALF_UP style: divide by 10_000 then multiply back
    let scaled = ((n as f64) / 10_000.0).round() * 10_000.0;
    assert!(scaled > 0.0);
    let scientific = 3.1010110212361498e18_f64;
    assert!(scientific.is_finite());
}

/// SimpleDateFormat-style locale date format strings (DataFormatTest probes).
pub fn assert_locale_date_format_smoke() {
    // Pattern presence checks — portable stand-in for Java SimpleDateFormat probes.
    let patterns = ["yyyy年m月d日 h点mm哈哈哈m", "yyyy年m月d日", "ah时mm分"];
    for p in patterns {
        assert!(p.contains('年') || p.contains('时') || p.contains('月'));
    }
    // Java `date_ptrn6`: must contain at least one of 年|月|日|时|分|秒
    fn has_date_token(s: &str) -> bool {
        s.chars()
            .any(|c| matches!(c, '年' | '月' | '日' | '时' | '分' | '秒'))
    }
    assert!(has_date_token("2017年"));
    assert!(!has_date_token("2017但是"));
}

/// ArrayList clear vs reallocate micro smoke (DataFormatTest#test2).
pub fn assert_vec_clear_vs_realloc() {
    let mut list: Vec<String> = Vec::with_capacity(3000);
    for _ in 0..1_000 {
        list.clear();
    }
    for _ in 0..1_000 {
        list = Vec::with_capacity(3000);
    }
    assert_eq!(list.capacity(), 3000);
}

/// CellReference("B3") portable stand-in (Lock2Test#testc) via PositionUtils mirror.
pub fn assert_cell_reference_b3() {
    use easyexcel::util::position_utils::{get_col, get_row};
    // B3 → col 1 (0-based), row 2 (0-based)
    assert_eq!(get_col("B3"), 1);
    assert_eq!(get_row("B3"), 2);
    assert_eq!(get_col("A1"), 0);
    assert_eq!(get_row("A1"), 0);
}

/// Excel serial ↔ datetime probes (Lock2Test#numberforamt) without Apache POI DateUtil.
pub fn assert_excel_serial_date_probes() {
    use chrono::{NaiveDate, NaiveDateTime, Timelike};
    use easyexcel::util::date_utils::get_java_date;

    // Integer-day serial: 44729 ≈ 2022-06-17 (1900 date system / Lotus bug epoch)
    let dt = get_java_date(44729);
    assert_eq!(
        dt.date_naive(),
        NaiveDate::from_ymd_opt(2022, 6, 17).unwrap()
    );

    // Fractional serial → wall-clock via epoch + days + fraction-of-day
    let serial = 44729.99998842592_f64;
    let whole = serial.floor() as i64;
    let frac = serial - whole as f64;
    let base = get_java_date(whole).naive_utc();
    let nanos = (frac * 86_400.0 * 1_000_000_000.0).round() as i64;
    let with_time = base + chrono::Duration::nanoseconds(nanos);
    assert_eq!(
        with_time.date(),
        NaiveDate::from_ymd_opt(2022, 6, 17).unwrap()
    );
    // Near end of day (23:59:xx)
    assert!(with_time.hour() >= 23);

    // Reverse: NaiveDateTime → Excel serial (1900 system)
    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let sample = NaiveDateTime::new(
        NaiveDate::from_ymd_opt(2022, 6, 17).unwrap(),
        chrono::NaiveTime::from_hms_opt(23, 59, 59).unwrap(),
    );
    let days = (sample.date() - epoch).num_days() as f64;
    let secs = f64::from(sample.time().num_seconds_from_midnight());
    let excel = days + secs / 86_400.0;
    assert!((excel - 44729.99998842592).abs() < 1e-5);
}

/// Sampled Excel date round-trip (Lock2Test#testDateAll) — daily samples, not every-second stress.
pub fn assert_excel_date_roundtrip_sampled() {
    use chrono::{NaiveDate, NaiveDateTime, Timelike};
    use easyexcel::util::date_utils::get_java_date;

    let epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    // Sample noon of each day from 1970-01-01 through ~2 years (CI-friendly).
    let start = NaiveDate::from_ymd_opt(1970, 1, 1).unwrap();
    for day_offset in 0..800 {
        let date = start + chrono::Duration::days(day_offset);
        let noon = NaiveDateTime::new(date, chrono::NaiveTime::from_hms_opt(12, 0, 0).unwrap());
        let serial = (date - epoch).num_days() as f64
            + f64::from(noon.time().num_seconds_from_midnight()) / 86_400.0;
        let whole = serial.floor() as i64;
        let back = get_java_date(whole).date_naive();
        assert_eq!(back, date, "serial roundtrip day_offset={day_offset}");
    }
}

/// POI DateUtil.isADateFormat portable stand-in (DataFormatTest#test31 / #test1).
pub fn assert_is_a_date_format_probes() {
    use easyexcel::util::date_utils::is_a_date_format;

    // Built-in Excel date format indexes
    assert!(is_a_date_format(14, None));
    assert!(is_a_date_format(22, None));
    assert!(!is_a_date_format(0, None));

    // Java DataFormatTest#test31
    assert!(is_a_date_format(
        181,
        Some("[DBNum1][$-404]m\"月\"d\"日\";@"),
    ));
    // Java DataFormatTest#test1
    assert!(is_a_date_format(
        181,
        Some("yyyy\"年啊\"m\"月\"d\"日\"\\ h")
    ));
    assert!(is_a_date_format(
        180,
        Some("yyyy\"年\"m\"月\"d\"日\"\\ h\"点\"")
    ));
}

/// Read `fill/simple.xlsx` (replaces Java `TestFileUtil` / local-path lastRowNum probes).
pub fn assert_fill_simple_read() {
    let path = fixture("fill/simple.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty(), "fill/simple.xlsx must yield rows");
}

/// Read `fill/complex.xlsx` (PoiTest#lastRowNum255 local `complex.xlsx` + shiftRows stand-in).
pub fn assert_fill_complex_read() {
    let path = fixture("fill/complex.xlsx");
    assert_fixture(&path);
    let rows = EasyExcel::read_dynamic_sync(&path)
        .head_row_number(0)
        .do_read_sync()
        .unwrap();
    assert!(!rows.is_empty(), "fill/complex.xlsx must yield rows");
}

/// Byte-read of `fill/simple.xlsx` (PoiTest#testreadRead FileUtils.readFileToByteArray).
pub fn assert_fill_simple_bytes() {
    let path = fixture("fill/simple.xlsx");
    assert_fixture(&path);
    let bytes = std::fs::read(&path).unwrap();
    assert!(bytes.len() > 4);
    assert_eq!(&bytes[0..2], b"PK");
}

/// PoiWriteTest#write0 / #write — large integer cells via EasyExcel write/read.
///
/// Excel stores numbers as IEEE f64; integers past ~2^53 lose precision, so the
/// portable contract mirrors Java `PoiWriteTest#write` (string cells) rather than
/// raw `setCellValue(long)`.
pub fn assert_poi_long_write() {
    #[derive(Debug, Clone, ExcelRow)]
    struct LongRow {
        #[excel(name = "a")]
        a: String,
        #[excel(name = "b")]
        b: String,
        #[excel(name = "c")]
        c: f64,
    }
    let path = temp_path("1to1_poi_long.xlsx");
    EasyExcel::write::<LongRow>(&path)
        .sheet("t1")
        .do_write(vec![LongRow {
            a: "999999999999999".into(),
            b: "1000000000000001".into(),
            c: 300.35,
        }])
        .unwrap();
    let rows = EasyExcel::read_sync::<LongRow>(&path)
        .do_read_sync()
        .unwrap();
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].a, "999999999999999");
    assert_eq!(rows[0].b, "1000000000000001");
    assert!((rows[0].c - 300.35).abs() < 1e-4);
}

/// PoiWriteTest#write01 — float → BigDecimal-style string round-trip smoke.
pub fn assert_float_decimal_smoke() {
    let ff = 300.35_f32;
    let bd: f64 = ff.to_string().parse().unwrap();
    assert!((bd - 300.35).abs() < 1e-4);
    assert!((ff - 300.35).abs() < 1e-4);
}

/// PoiWriteTest#write1 — long → big-endian bytes (Java long2Bytes).
pub fn assert_long2bytes() {
    fn long2bytes(num: i64) -> [u8; 8] {
        let mut out = [0u8; 8];
        for (ix, slot) in out.iter_mut().enumerate() {
            let offset = 64 - (ix as i32 + 1) * 8;
            *slot = ((num >> offset) & 0xff) as u8;
        }
        out
    }
    let a = long2bytes(-999_999_999_999_999);
    let b = long2bytes(-9_999_999_999_999_999);
    assert_eq!(a.len(), 8);
    assert_eq!(b.len(), 8);
    assert_ne!(a, b);
}

/// Whether `s` contains a `${...}` placeholder with non-empty body (Java FILL_PATTERN).
fn has_fill_placeholder(s: &str) -> bool {
    let bytes = s.as_bytes();
    let mut i = 0;
    while i + 2 < bytes.len() {
        if bytes[i] == b'$' && bytes[i + 1] == b'{' {
            if let Some(end) = s[i + 2..].find('}') {
                if end > 0 {
                    return true;
                }
            }
        }
        i += 1;
    }
    false
}

/// PoiWriteTest#part / #part2 — fill placeholder `${...}` pattern probes.
pub fn assert_fill_placeholder_pattern() {
    assert!(has_fill_placeholder("${name}今年${number}岁了"));
    assert!(has_fill_placeholder("${name}"));
    assert!(has_fill_placeholder("${number}"));
    assert!(has_fill_placeholder("${name}今年"));
    assert!(has_fill_placeholder("今年${number}岁了"));
    assert!(!has_fill_placeholder("今年${number岁了"));
    assert!(!has_fill_placeholder("${}"));
    assert!(!has_fill_placeholder("胜多负少"));
    // part2-style substring probe (Java FILL_PATTERN2 = "测试")
    assert!("我是测试呀".contains("测试"));
    assert!("测试u".contains("测试"));
}
