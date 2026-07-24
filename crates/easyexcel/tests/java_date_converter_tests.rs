//! End-to-end parity tests for distinct Java Date and LocalDateTime converter keys.

use chrono::{NaiveDate, NaiveDateTime};
use easyexcel::{EasyExcel, ExcelRow, JavaDate, Result};
use tempfile::tempdir;

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct DistinctDateRow {
    #[excel(name = "Java Date", index = 0)]
    java_date: JavaDate,
    #[excel(name = "Local Date Time", index = 1)]
    local_date_time: NaiveDateTime,
}

#[test]
fn java_date_and_local_date_time_default_converters_coexist_end_to_end() -> Result<()> {
    let directory = tempdir()?;
    let path = directory.path().join("distinct-date-keys.xlsx");
    let datetime = NaiveDate::from_ymd_opt(2026, 7, 24)
        .unwrap()
        .and_hms_opt(12, 34, 56)
        .unwrap();
    let rows = vec![DistinctDateRow {
        java_date: JavaDate::from(datetime),
        local_date_time: datetime,
    }];

    EasyExcel::write::<DistinctDateRow>(&path)
        .sheet("Dates")
        .do_write(rows.clone())?;
    assert_eq!(
        EasyExcel::read_sync::<DistinctDateRow>(&path).do_read_sync()?,
        rows
    );
    Ok(())
}
