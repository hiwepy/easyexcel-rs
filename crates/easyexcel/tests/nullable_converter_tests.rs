use easyexcel::{
    CellValue, Converter, EasyExcel, ExcelRow, NullableObjectConverter, Result, WriteCellData,
    WriteConverterContext,
};
use tempfile::tempdir;

#[derive(Debug, Clone, PartialEq, Eq, ExcelRow)]
struct OptionalTextRow {
    #[excel(name = "Value", index = 0)]
    value: Option<String>,
}

#[derive(Debug, Clone, Copy)]
struct OrdinaryOptionConverter;

impl Converter<Option<String>> for OrdinaryOptionConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, Option<String>>,
    ) -> Result<WriteCellData> {
        Ok(WriteCellData::new(CellValue::String(
            context
                .value()
                .as_deref()
                .unwrap_or("ordinary-null")
                .to_owned(),
        )))
    }
}

#[derive(Debug, Clone, Copy)]
struct NullableOptionConverter;

impl Converter<Option<String>> for NullableOptionConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, Option<String>>,
    ) -> Result<WriteCellData> {
        Ok(WriteCellData::new(CellValue::String(
            context
                .value()
                .as_deref()
                .unwrap_or("nullable-null")
                .to_owned(),
        )))
    }
}

impl NullableObjectConverter<Option<String>> for NullableOptionConverter {}

#[test]
fn derive_and_facade_apply_java_nullable_converter_gate_end_to_end() -> Result<()> {
    let directory = tempdir()?;
    let ordinary_path = directory.path().join("ordinary.xlsx");
    let nullable_path = directory.path().join("nullable.xlsx");
    let rows = vec![OptionalTextRow { value: None }];

    EasyExcel::write::<OptionalTextRow>(&ordinary_path)
        .register_converter::<Option<String>, _>(OrdinaryOptionConverter)
        .sheet("Data")
        .do_write(rows.clone())?;
    assert!(
        EasyExcel::read_sync::<OptionalTextRow>(&ordinary_path)
            .do_read_sync()?
            .is_empty(),
        "the ordinary converter is skipped and the all-empty physical row is omitted"
    );

    EasyExcel::write::<OptionalTextRow>(&nullable_path)
        .register_nullable_converter::<Option<String>, _>(NullableOptionConverter)
        .sheet("Data")
        .do_write(rows)?;
    assert_eq!(
        EasyExcel::read_sync::<OptionalTextRow>(&nullable_path).do_read_sync()?,
        [OptionalTextRow {
            value: Some("nullable-null".to_owned())
        }]
    );
    Ok(())
}
