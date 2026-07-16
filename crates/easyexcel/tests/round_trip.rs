//! End-to-end compatibility tests for the public facade.

use std::cell::RefCell;
use std::rc::Rc;

use chrono::NaiveDate;
use easyexcel::{
    AnalysisContext, CellValue, Converter, EasyExcel, ExcelRow, PageReadListener,
    ReadConverterContext, ReadListener, Result, WriteConverterContext,
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
