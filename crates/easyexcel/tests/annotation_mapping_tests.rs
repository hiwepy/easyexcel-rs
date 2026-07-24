//! Java annotation semantics exercised through Rust derive metadata and real XLSX I/O.

use std::str::FromStr;

use bigdecimal::BigDecimal;
use chrono::NaiveDate;
use easyexcel::{EasyExcel, ExcelRow, NumberRoundingMode, Result};
use tempfile::tempdir;

#[derive(Debug, PartialEq, ExcelRow)]
#[excel(ignore_unannotated)]
struct AnnotationModel {
    unannotated: String,
    #[excel(name = "Amount", number_format = "0.0", rounding_mode = "HALF_DOWN")]
    amount: BigDecimal,
    #[excel(name = "Date", format = "%Y-%m-%d", use_1904_windowing = true)]
    date: NaiveDate,
    #[excel(ignore, name = "Ignored")]
    ignored: String,
}

#[test]
fn derive_applies_ignore_and_format_annotations_to_real_model_mapping() -> Result<()> {
    let schema = AnnotationModel::schema();
    assert_eq!(schema.len(), 2);
    assert_eq!(schema[0].field, "amount");
    assert_eq!(schema[0].name, "Amount");
    assert_eq!(schema[0].format, Some("0.0"));
    assert_eq!(
        schema[0].number_rounding_mode,
        Some(NumberRoundingMode::HalfDown)
    );
    assert_eq!(schema[1].field, "date");
    assert_eq!(schema[1].format, Some("%Y-%m-%d"));
    assert_eq!(schema[1].use_1904_windowing, Some(true));

    let directory = tempdir()?;
    let path = directory.path().join("annotation-mapping.xlsx");
    let expected = AnnotationModel {
        unannotated: String::new(),
        amount: BigDecimal::from_str("12.5")
            .map_err(|error| easyexcel::ExcelError::Format(error.to_string()))?,
        date: NaiveDate::from_ymd_opt(2026, 7, 24).expect("valid date"),
        ignored: String::new(),
    };
    EasyExcel::write::<AnnotationModel>(&path).do_write([expected])?;

    let rows = EasyExcel::read_sync::<AnnotationModel>(&path).do_read_sync()?;
    assert_eq!(rows.len(), 1);
    assert_eq!(rows[0].unannotated, "");
    assert_eq!(rows[0].ignored, "");
    assert_eq!(
        rows[0].amount,
        BigDecimal::from(125_i32) / BigDecimal::from(10_i32)
    );
    assert_eq!(
        rows[0].date,
        NaiveDate::from_ymd_opt(2026, 7, 24).expect("valid date")
    );
    Ok(())
}
