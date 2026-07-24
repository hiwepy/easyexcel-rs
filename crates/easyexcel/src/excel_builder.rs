//! Facade wiring for Java-style [`ExcelBuilderImpl::fill`].
//!
//! `easyexcel-writer` keeps a thin [`WriteFillExecutor`] hook; this module
//! connects it to `easyexcel-template` without introducing a crate cycle.

use std::any::Any;
use std::path::PathBuf;

use easyexcel_core::{DynamicRow, Result};
use easyexcel_template::create_builder_fill_executor;
use easyexcel_writer::BuilderFillConfig;
use easyexcel_writer::{ExcelBuilder, ExcelBuilderImpl, ExcelWriter, WriteSheet};

/// Creates an [`ExcelBuilderImpl`] from a stateful writer without fill wiring.
///
/// Use [`fill_builder_from_writer`] when the builder will call [`ExcelBuilder::fill`].
#[must_use]
pub fn builder_from_writer(writer: ExcelWriter) -> ExcelBuilderImpl {
    let path = writer.output_path().to_path_buf();
    ExcelBuilderImpl::new(writer, path)
}

/// Creates an [`ExcelBuilderImpl`] and wires template fill when configured.
///
/// Mirrors Java `new ExcelBuilderImpl(WriteWorkbook)` where
/// `WriteWorkbook.templateInputStream` is non-null.
///
/// # Errors
///
/// Returns an I/O or OOXML error when the configured template cannot be loaded.
pub fn fill_builder_from_writer(writer: ExcelWriter) -> Result<ExcelBuilderImpl> {
    let output = writer.output_path().to_path_buf();
    let template_file = writer.template_file().map(PathBuf::from);
    let template_bytes = writer.template_bytes().map(<[u8]>::to_vec);
    let mut builder = ExcelBuilderImpl::new(writer, output.clone());
    if template_file.is_some() || template_bytes.is_some() {
        let executor = create_builder_fill_executor(template_file, template_bytes, output)?;
        builder.set_fill_executor(executor);
    }
    Ok(builder)
}

/// Wires template fill into an existing builder when a template is configured.
///
/// # Errors
///
/// Returns an I/O or OOXML error when the configured template cannot be loaded.
pub fn wire_template_fill(builder: &mut ExcelBuilderImpl) -> Result<()> {
    if builder.has_fill_executor() {
        return Ok(());
    }
    let writer = builder.writer_mut();
    if !writer.has_template_configured() {
        return Ok(());
    }
    let output = writer.output_path().to_path_buf();
    let template_file = writer.template_file().map(PathBuf::from);
    let template_bytes = writer.template_bytes().map(<[u8]>::to_vec);
    let executor = create_builder_fill_executor(template_file, template_bytes, output)?;
    builder.set_fill_executor(executor);
    Ok(())
}

/// Executes one Java-style `doFill` through [`ExcelBuilderImpl`].
///
/// Mirrors `EasyExcel.write(file).withTemplate(template).sheet().doFill(data)`.
///
/// # Errors
///
/// Returns template, fill, or output errors from the wired builder path.
pub fn do_fill_template(
    writer: ExcelWriter,
    data: &dyn Any,
    sheet: &WriteSheet<DynamicRow>,
) -> Result<()> {
    do_fill_template_with_config(writer, data, BuilderFillConfig::default(), sheet)
}

/// Executes `doFill` with an explicit builder [`FillConfig`].
///
/// # Errors
///
/// Returns template, fill, or output errors from the wired builder path.
pub fn do_fill_template_with_config(
    writer: ExcelWriter,
    data: &dyn Any,
    fill_config: BuilderFillConfig,
    sheet: &WriteSheet<DynamicRow>,
) -> Result<()> {
    let mut builder = fill_builder_from_writer(writer)?;
    builder.fill(data, fill_config, sheet)?;
    builder.finish(false)
}

#[cfg(test)]
mod tests {
    use super::*;
    use easyexcel_core::DynamicValue;
    use easyexcel_template::{FillConfig, FillWrapper, TemplateData};
    use std::collections::BTreeMap;
    use tempfile::tempdir;

    #[test]
    fn fill_builder_from_writer_delegates_scalar_fill() -> Result<()> {
        let directory = tempdir()?;
        let template = directory.path().join("template.xlsx");
        let output = directory.path().join("filled.xlsx");
        EasyExcel::write::<DynamicRow>(&template)
            .need_head(false)
            .do_write([DynamicRow::new({
                let mut cells = BTreeMap::new();
                cells.insert(0, DynamicValue::String("{name}".to_owned()));
                cells
            })])?;

        let writer = EasyExcel::write::<DynamicRow>(&output)
            .with_template(&template)
            .need_head(false)
            .build();
        let sheet = WriteSheet::<DynamicRow>::new("Sheet1");
        do_fill_template_with_config(
            writer,
            &TemplateData::new().with("name", "builder-fill"),
            BuilderFillConfig::default(),
            &sheet,
        )?;
        assert!(output.exists());
        let rows = EasyExcel::read_dynamic_sync(&output)
            .head_row_number(0)
            .do_read_sync()?;
        assert!(rows.iter().any(|row| {
            row.values().values().any(|value| {
                matches!(value, DynamicValue::String(text) if text.contains("builder-fill"))
            })
        }));
        Ok(())
    }

    #[test]
    fn facade_do_fill_accepts_collection_config_and_supplier() -> Result<()> {
        let directory = tempdir()?;
        let template = directory.path().join("list-template.xlsx");
        let list_output = directory.path().join("list-filled.xlsx");
        let supplier_output = directory.path().join("supplier-filled.xlsx");
        EasyExcel::write::<DynamicRow>(&template)
            .need_head(false)
            .do_write([DynamicRow::new({
                let mut cells = BTreeMap::new();
                cells.insert(0, DynamicValue::String("{.name}".to_owned()));
                cells.insert(1, DynamicValue::String("{title}".to_owned()));
                cells
            })])?;

        EasyExcel::write::<DynamicRow>(&list_output)
            .with_template(&template)
            .need_head(false)
            .do_fill_with_config(
                &FillWrapper::new([
                    TemplateData::new().with("name", "A"),
                    TemplateData::new().with("name", "B"),
                ]),
                FillConfig::new().auto_style(false),
            )?;

        let list_rows = EasyExcel::read_dynamic_sync(&list_output)
            .head_row_number(0)
            .do_read_sync()?;
        let rendered = list_rows
            .iter()
            .flat_map(|row| row.values().values())
            .filter_map(|value| match value {
                DynamicValue::String(text) => Some(text.as_str()),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert!(rendered.contains(&"A"));
        assert!(rendered.contains(&"B"));

        EasyExcel::write::<DynamicRow>(&supplier_output)
            .with_template(&template)
            .need_head(false)
            .do_fill_with(|| TemplateData::new().with("title", "Supplier"))?;
        let supplier_rows = EasyExcel::read_dynamic_sync(&supplier_output)
            .head_row_number(0)
            .do_read_sync()?;
        assert!(supplier_rows.iter().any(|row| {
            row.values()
                .values()
                .any(|value| matches!(value, DynamicValue::String(text) if text == "Supplier"))
        }));
        Ok(())
    }

    use crate::EasyExcel;
}
