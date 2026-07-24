//! Bridges [`WriteFillExecutor`] to [`ExcelTemplateWriter`].
//!
//! Keeps template fill logic out of `easyexcel-writer` while letting
//! `ExcelBuilderImpl.fill` delegate to the same engine as
//! `EasyExcel::template_writer`.

use std::any::Any;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use easyexcel_core::{
    ExcelError, Result, WriteDirection, WriteFillConfig, WriteFillExecutor, WriteFillSheet,
};

use crate::{
    ExcelTemplateWriter, FillConfig, FillDirection, FillWrapper, TemplateData, TemplateSheet,
};

/// Stateful template fill executor for [`easyexcel_writer::ExcelBuilderImpl`].
///
/// Mirrors Java `ExcelWriteFillExecutor` backed by the same loaded XLSX
/// package as [`ExcelTemplateWriter`].
pub struct BuilderFillExecutor {
    inner: ExcelTemplateWriter<'static>,
}

impl BuilderFillExecutor {
    /// Loads a template from path or bytes and prepares fill against `output`.
    ///
    /// # Errors
    ///
    /// Returns I/O or OOXML package errors when the template cannot be read.
    pub fn new(
        template_file: Option<PathBuf>,
        template_bytes: Option<Vec<u8>>,
        output: PathBuf,
    ) -> Result<Self> {
        let inner = if let Some(path) = template_file {
            ExcelTemplateWriter::new(path, output)?
        } else if let Some(bytes) = template_bytes {
            ExcelTemplateWriter::from_reader(Cursor::new(bytes), output)?
        } else {
            return Err(ExcelError::Unsupported(
                "with_template requires a template file or template bytes".to_owned(),
            ));
        };
        Ok(Self { inner })
    }

    /// Loads a template file and writes to an existing path.
    ///
    /// # Errors
    ///
    /// Returns I/O or OOXML package errors when the template cannot be read.
    pub fn from_template_path(
        template: impl AsRef<Path>,
        output: impl Into<PathBuf>,
    ) -> Result<Self> {
        Ok(Self {
            inner: ExcelTemplateWriter::new(template, output)?,
        })
    }
}

/// Creates a boxed fill executor for facade wiring into [`ExcelBuilderImpl`](easyexcel_writer::ExcelBuilderImpl).
///
/// # Errors
///
/// Returns I/O or OOXML package errors when the template cannot be read.
pub fn create_builder_fill_executor(
    template_file: Option<PathBuf>,
    template_bytes: Option<Vec<u8>>,
    output: PathBuf,
) -> Result<Box<dyn WriteFillExecutor>> {
    Ok(Box::new(BuilderFillExecutor::new(
        template_file,
        template_bytes,
        output,
    )?))
}

impl WriteFillExecutor for BuilderFillExecutor {
    fn fill(
        &mut self,
        data: &dyn Any,
        fill_config: WriteFillConfig,
        sheet: WriteFillSheet,
    ) -> Result<()> {
        let template_sheet = to_template_sheet(&sheet);
        if let Some(scalar) = data.downcast_ref::<TemplateData>() {
            self.inner.fill_on_sheet(&template_sheet, scalar)?;
            return Ok(());
        }
        if let Some(collection) = data.downcast_ref::<FillWrapper>() {
            self.inner.fill_list_on_sheet(
                &template_sheet,
                collection,
                to_template_fill_config(fill_config),
            )?;
            return Ok(());
        }
        Err(ExcelError::Format(format!(
            "fill data must be TemplateData or FillWrapper, got {}",
            std::any::type_name_of_val(data)
        )))
    }

    fn finish(&mut self, _on_exception: bool) -> Result<()> {
        self.inner.finish()
    }
}

fn to_template_sheet(sheet: &WriteFillSheet) -> TemplateSheet {
    if let Some(index) = sheet.sheet_index {
        TemplateSheet::index(index)
    } else if sheet.sheet_name.chars().all(|ch| ch.is_ascii_digit())
        && let Ok(index) = sheet.sheet_name.parse::<usize>()
    {
        TemplateSheet::index(index)
    } else {
        TemplateSheet::name(sheet.sheet_name.clone())
    }
}

fn to_template_fill_config(config: WriteFillConfig) -> FillConfig {
    let mut fill_config = FillConfig::new()
        .force_new_row(config.force_new_row)
        .auto_style(config.auto_style);
    if let Some(direction) = config.direction {
        fill_config = fill_config.direction(match direction {
            WriteDirection::Vertical => FillDirection::Vertical,
            WriteDirection::Horizontal => FillDirection::Horizontal,
        });
    }
    fill_config
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn builder_fill_config_propagates_direction_force_row_and_auto_style() {
        let config = to_template_fill_config(WriteFillConfig {
            force_new_row: true,
            direction: Some(WriteDirection::Horizontal),
            auto_style: false,
        });

        assert_eq!(config.get_direction(), FillDirection::Horizontal);
        assert!(config.get_force_new_row());
        assert!(!config.get_auto_style());
    }
}
