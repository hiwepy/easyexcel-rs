use std::collections::HashMap;
use std::io;
use std::sync::Arc;

use chrono::{NaiveDate, NaiveDateTime};

use super::*;

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
        (CellValue::Date(date), "2026-07-17"),
        (CellValue::DateTime(datetime), "2026-07-17 12:34:56"),
    ];
    for (value, expected) in cases {
        assert_eq!(value.as_text(), expected);
    }
    assert!(CellValue::Empty.is_empty());
    assert!(!CellValue::Bool(false).is_empty());
}

#[test]
fn row_data_resolves_index_before_header_name() {
    let explicit = ExcelColumn::new("first", "Header", Some(1), 3, Some("0"));
    let named = ExcelColumn::new("second", "Header", None, i32::MAX, None);
    let missing = ExcelColumn::new("missing", "Missing", None, i32::MAX, None);
    let headers = Arc::new(HashMap::from([("Header".to_owned(), 0)]));
    let row = RowData::new(
        "Users",
        7,
        vec![CellValue::String("name".to_owned()), CellValue::Int(9)],
        headers,
    );

    assert_eq!(row.sheet_name(), "Users");
    assert_eq!(row.row_index(), 7);
    assert_eq!(row.cell(&explicit), Some(&CellValue::Int(9)));
    assert_eq!(
        row.cell(&named),
        Some(&CellValue::String("name".to_owned()))
    );
    assert_eq!(row.cell(&missing), None);
    assert_eq!(row.convert_context(&explicit).column_index, Some(1));
    assert_eq!(row.convert_context(&named).column_index, Some(0));
    assert_eq!(row.convert_context(&missing).column_index, None);
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
fn floating_point_types_support_numeric_and_string_cells() -> Result<()> {
    let context = context(None);
    for value in [
        f32::from_excel_cell(Some(&CellValue::Int(2)), &context)?,
        f32::from_excel_cell(Some(&CellValue::Float(2.0)), &context)?,
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

#[test]
fn analysis_context_exposes_sheet_row_and_batch_coordinates() {
    let context = AnalysisContext::new("Users", 3, 9);
    assert_eq!(context.sheet_name(), "Users");
    assert_eq!(context.sheet_no(), 3);
    assert_eq!(context.row_index(), 9);
    assert_eq!(context.batch_index(), 0);
    assert_eq!(context.with_batch_index(4).batch_index(), 4);
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
    let mut defaults = DefaultListener;
    assert_eq!(defaults.on_exception(&error, &context), ErrorAction::Stop);
    defaults.invoke_head(&HashMap::new(), &context)?;
    defaults.invoke(1, &context)?;
    defaults.do_after_all_analysed(&context)?;
    assert!(defaults.has_next(&context));

    let mut listener: Box<dyn ReadListener<i32>> = Box::new(RecordingListener::default());
    assert_eq!(
        listener.on_exception(&error, &context),
        ErrorAction::Continue
    );
    listener.invoke_head(&HashMap::new(), &context)?;
    listener.invoke(1, &context)?;
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
