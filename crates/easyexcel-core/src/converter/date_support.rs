use std::borrow::Cow;

use chrono::{NaiveDate, NaiveDateTime, Timelike};

use crate::util::work_book_util::fill_data_format;
use crate::{
    CellValue, ExcelError, FromExcelCell, ReadConverterContext, WriteCellData,
    WriteConverterContext,
};

pub(crate) const DEFAULT_DATE_FORMAT: &str = "yyyy-MM-dd";
pub(crate) const DEFAULT_DATETIME_FORMAT: &str = "yyyy-MM-dd HH:mm:ss";

pub(crate) fn read_date(context: &ReadConverterContext<'_>) -> Result<NaiveDate, ExcelError> {
    if let Some(CellValue::String(value)) = context.cell() {
        let patterns = context
            .convert_context()
            .format
            .map_or_else(|| vec!["%Y-%m-%d", "%Y/%m/%d"], |pattern| vec![pattern]);
        return patterns
            .into_iter()
            .find_map(|pattern| {
                let pattern = chrono_pattern(pattern);
                NaiveDate::parse_from_str(value, &pattern).ok()
            })
            .ok_or_else(|| {
                context
                    .convert_context()
                    .invalid(context.cell().expect("string cell exists"), "NaiveDate")
            });
    }
    NaiveDate::from_excel_cell(context.cell(), context.convert_context())
}

pub(crate) fn read_datetime(
    context: &ReadConverterContext<'_>,
) -> Result<NaiveDateTime, ExcelError> {
    if let Some(CellValue::String(value)) = context.cell() {
        let patterns = context.convert_context().format.map_or_else(
            || {
                vec![
                    "%Y%m%d%H%M%S",
                    "%Y-%m-%d %H:%M",
                    "%Y/%m/%d %H:%M",
                    "%Y%m%d %H:%M:%S",
                    "%Y-%m-%d %H:%M:%S",
                    "%Y/%m/%d %H:%M:%S",
                ]
            },
            |pattern| vec![pattern],
        );
        return patterns
            .into_iter()
            .find_map(|pattern| {
                let pattern = chrono_pattern(pattern);
                NaiveDateTime::parse_from_str(value, &pattern).ok()
            })
            .ok_or_else(|| {
                context
                    .convert_context()
                    .invalid(context.cell().expect("string cell exists"), "NaiveDateTime")
            });
    }
    NaiveDateTime::from_excel_cell(context.cell(), context.convert_context())
}

pub(crate) fn write_date_value(
    value: NaiveDate,
    context: &WriteConverterContext<'_, NaiveDate>,
) -> WriteCellData {
    let mut cell = WriteCellData::new(CellValue::Date(value));
    fill_data_format(
        &mut cell,
        context.convert_context().format,
        DEFAULT_DATE_FORMAT,
    );
    cell
}

pub(crate) fn write_datetime_value<T>(
    value: NaiveDateTime,
    context: &WriteConverterContext<'_, T>,
) -> WriteCellData {
    let mut cell = WriteCellData::new(CellValue::DateTime(value));
    fill_data_format(
        &mut cell,
        context.convert_context().format,
        DEFAULT_DATETIME_FORMAT,
    );
    cell
}

pub(crate) fn write_date_string(
    value: NaiveDate,
    context: &WriteConverterContext<'_, NaiveDate>,
) -> WriteCellData {
    let pattern = context.convert_context().format.unwrap_or("%Y-%m-%d");
    WriteCellData::from_string(value.format(&chrono_pattern(pattern)).to_string())
}

pub(crate) fn write_datetime_string<T>(
    value: NaiveDateTime,
    context: &WriteConverterContext<'_, T>,
) -> WriteCellData {
    let pattern = context
        .convert_context()
        .format
        .unwrap_or("%Y-%m-%d %H:%M:%S");
    WriteCellData::from_string(value.format(&chrono_pattern(pattern)).to_string())
}

fn chrono_pattern(pattern: &str) -> Cow<'_, str> {
    if pattern.contains('%') {
        return Cow::Borrowed(pattern);
    }
    Cow::Owned(
        pattern
            .replace("yyyy", "%Y")
            .replace("SSS", "%.3f")
            .replace("MM", "%m")
            .replace("dd", "%d")
            .replace("HH", "%H")
            .replace("mm", "%M")
            .replace("ss", "%S"),
    )
}

pub(crate) fn format_number_as_datetime_string(
    context: &ReadConverterContext<'_>,
    pattern: &str,
) -> Result<String, ExcelError> {
    let value = NaiveDateTime::from_excel_cell(context.cell(), context.convert_context())?;
    Ok(value.format(&chrono_pattern(pattern)).to_string())
}

pub(crate) fn date_to_excel_serial(value: NaiveDate, use_1904_windowing: bool) -> f64 {
    let epoch = if use_1904_windowing {
        NaiveDate::from_ymd_opt(1904, 1, 1).expect("valid Excel epoch")
    } else if value < NaiveDate::from_ymd_opt(1900, 3, 1).expect("valid Excel boundary") {
        NaiveDate::from_ymd_opt(1899, 12, 31).expect("valid Excel epoch")
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).expect("valid Excel epoch")
    };
    (value - epoch).num_days() as f64
}

pub(crate) fn datetime_to_excel_serial(value: NaiveDateTime, use_1904_windowing: bool) -> f64 {
    let seconds = f64::from(value.time().num_seconds_from_midnight())
        + f64::from(value.time().nanosecond()) / 1_000_000_000.0;
    date_to_excel_serial(value.date(), use_1904_windowing) + seconds / 86_400.0
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::converter::date::date_date_converter::DateDateConverter;
    use crate::converter::date::date_number_converter::DateNumberConverter;
    use crate::converter::date::date_string_converter::DateStringConverter;
    use crate::converter::localdate::local_date_date_converter::LocalDateDateConverter;
    use crate::converter::localdate::local_date_number_converter::LocalDateNumberConverter;
    use crate::converter::localdate::local_date_string_converter::LocalDateStringConverter;
    use crate::converter::localdatetime::local_date_time_date_converter::LocalDateTimeDateConverter;
    use crate::converter::localdatetime::local_date_time_number_converter::LocalDateTimeNumberConverter;
    use crate::converter::localdatetime::local_date_time_string_converter::LocalDateTimeStringConverter;
    use crate::{CellDataType, ConvertContext, Converter, ExcelColumn, JavaDate};

    const COLUMN: ExcelColumn = ExcelColumn::new("value", "Value", Some(0), 0, None);

    fn context(format: Option<&'static str>, use_1904_windowing: bool) -> ConvertContext {
        ConvertContext {
            sheet_name: "Sheet1".to_owned(),
            row_index: 1,
            column_index: Some(0),
            field: "value",
            format,
            use_1904_windowing,
        }
    }

    #[test]
    fn number_converters_are_real_bidirectional_java_equivalents() {
        let context_1900 = context(None, false);
        let serial = CellValue::Float(1.5);
        let read = ReadConverterContext::new(Some(&serial), &COLUMN, &context_1900);

        let date = LocalDateNumberConverter
            .convert_to_rust_data(&read)
            .expect("local date from number");
        assert_eq!(date, NaiveDate::from_ymd_opt(1900, 1, 1).unwrap());

        let date_time = LocalDateTimeNumberConverter
            .convert_to_rust_data(&read)
            .expect("local datetime from number");
        assert_eq!(
            date_time,
            NaiveDate::from_ymd_opt(1900, 1, 1)
                .unwrap()
                .and_hms_opt(12, 0, 0)
                .unwrap()
        );
        assert_eq!(
            DateNumberConverter
                .convert_to_rust_data(&read)
                .expect("java date equivalent")
                .naive_local(),
            date_time
        );
        let datetime_write = WriteConverterContext::new(&date_time, &COLUMN, &context_1900);
        let java_date = JavaDate::from(date_time);
        let java_date_write = WriteConverterContext::new(&java_date, &COLUMN, &context_1900);
        assert_eq!(
            DateNumberConverter
                .convert_to_excel_data(&java_date_write)
                .expect("java date to number")
                .value(),
            &CellValue::Float(1.5)
        );
        assert_eq!(
            LocalDateTimeNumberConverter
                .convert_to_excel_data(&datetime_write)
                .expect("local datetime to number")
                .value(),
            &CellValue::Float(1.5)
        );

        let date_write = WriteConverterContext::new(&date, &COLUMN, &context_1900);
        assert_eq!(
            LocalDateNumberConverter
                .convert_to_excel_data(&date_write)
                .unwrap()
                .value(),
            &CellValue::Float(1.0)
        );

        let context_1904 = context(None, true);
        let epoch_1904 = NaiveDate::from_ymd_opt(1904, 1, 1).unwrap();
        let write_1904 = WriteConverterContext::new(&epoch_1904, &COLUMN, &context_1904);
        assert_eq!(
            LocalDateNumberConverter
                .convert_to_excel_data(&write_1904)
                .unwrap()
                .value(),
            &CellValue::Float(0.0)
        );
        assert_eq!(
            <LocalDateNumberConverter as Converter<NaiveDate>>::support_excel_type(
                &LocalDateNumberConverter
            ),
            CellDataType::Number
        );
    }

    #[test]
    fn string_converters_honor_field_format_and_reject_invalid_input() {
        let date_context = context(Some("dd/MM/yyyy"), false);
        let date_cell = CellValue::String("31/12/2025".to_owned());
        let date_read = ReadConverterContext::new(Some(&date_cell), &COLUMN, &date_context);
        let date = LocalDateStringConverter
            .convert_to_rust_data(&date_read)
            .expect("formatted local date");
        assert_eq!(date, NaiveDate::from_ymd_opt(2025, 12, 31).unwrap());
        let date_write = WriteConverterContext::new(&date, &COLUMN, &date_context);
        assert_eq!(
            LocalDateStringConverter
                .convert_to_excel_data(&date_write)
                .unwrap()
                .value(),
            &CellValue::String("31/12/2025".to_owned())
        );

        let datetime_context = context(Some("yyyy-MM-dd HH:mm"), false);
        let datetime_cell = CellValue::String("2025-12-31 23:45".to_owned());
        let datetime_read =
            ReadConverterContext::new(Some(&datetime_cell), &COLUMN, &datetime_context);
        let datetime = LocalDateTimeStringConverter
            .convert_to_rust_data(&datetime_read)
            .expect("formatted local datetime");
        assert_eq!(
            DateStringConverter
                .convert_to_rust_data(&datetime_read)
                .expect("java date equivalent")
                .naive_local(),
            datetime
        );
        let datetime_write = WriteConverterContext::new(&datetime, &COLUMN, &datetime_context);
        let java_date = JavaDate::from(datetime);
        let java_date_write = WriteConverterContext::new(&java_date, &COLUMN, &datetime_context);
        assert_eq!(
            LocalDateTimeStringConverter
                .convert_to_excel_data(&datetime_write)
                .unwrap()
                .value(),
            &CellValue::String("2025-12-31 23:45".to_owned())
        );
        assert_eq!(
            DateStringConverter
                .convert_to_excel_data(&java_date_write)
                .unwrap()
                .value(),
            &CellValue::String("2025-12-31 23:45".to_owned())
        );

        let automatic_context = context(None, false);
        let automatic_cell = CellValue::String("2025/12/31 23:45:01".to_owned());
        let automatic_read =
            ReadConverterContext::new(Some(&automatic_cell), &COLUMN, &automatic_context);
        assert_eq!(
            LocalDateTimeStringConverter
                .convert_to_rust_data(&automatic_read)
                .expect("Java switchDateFormat equivalent"),
            NaiveDate::from_ymd_opt(2025, 12, 31)
                .unwrap()
                .and_hms_opt(23, 45, 1)
                .unwrap()
        );

        let invalid = CellValue::String("not-a-date".to_owned());
        let invalid_read = ReadConverterContext::new(Some(&invalid), &COLUMN, &date_context);
        assert!(
            LocalDateStringConverter
                .convert_to_rust_data(&invalid_read)
                .is_err()
        );
    }

    #[test]
    fn date_cell_converters_attach_java_default_data_formats() {
        let date = NaiveDate::from_ymd_opt(2025, 1, 2).unwrap();
        let date_context = context(None, false);
        let date_write = WriteConverterContext::new(&date, &COLUMN, &date_context);
        let date_cell = LocalDateDateConverter
            .convert_to_excel_data(&date_write)
            .expect("date cell");
        assert_eq!(date_cell.value(), &CellValue::Date(date));
        assert_eq!(
            date_cell.data_format_data().and_then(|data| data.format()),
            Some(DEFAULT_DATE_FORMAT)
        );

        let datetime = date.and_hms_opt(3, 4, 5).unwrap();
        let datetime_write = WriteConverterContext::new(&datetime, &COLUMN, &date_context);
        let java_date = JavaDate::from(datetime);
        let java_date_write = WriteConverterContext::new(&java_date, &COLUMN, &date_context);
        let java_cell = DateDateConverter
            .convert_to_excel_data(&java_date_write)
            .expect("java date cell");
        let local_cell = LocalDateTimeDateConverter
            .convert_to_excel_data(&datetime_write)
            .expect("local datetime cell");
        for cell in [java_cell, local_cell] {
            assert_eq!(cell.value(), &CellValue::DateTime(datetime));
            assert_eq!(
                cell.data_format_data().and_then(|data| data.format()),
                Some(DEFAULT_DATETIME_FORMAT)
            );
        }
    }
}
