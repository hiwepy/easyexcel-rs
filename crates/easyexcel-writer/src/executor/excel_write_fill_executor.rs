//! Mirrors Java `com.alibaba.excel.write.executor.ExcelWriteFillExecutor`.

use std::any::Any;

use crate::executor::abstract_excel_write_executor::AbstractExcelWriteExecutor;
use easyexcel_core::{
    ExcelError, Result, WriteContext, WriteFillConfig, WriteFillExecutor, WriteFillSheet,
};

/// Mirrors Java `ExcelWriteFillExecutor extends AbstractExcelWriteExecutor`.
///
/// The Java class owns the template-analysis and collection cursor state.
/// Rust keeps that state in the pluggable [`WriteFillExecutor`] implemented by
/// `easyexcel-template`; this executor is the production adapter used by
/// `ExcelBuilderImpl`, not a marker-only compatibility type.
pub struct ExcelWriteFillExecutor<'a> {
    inner: AbstractExcelWriteExecutor<'a>,
    delegate: Option<&'a mut dyn WriteFillExecutor>,
}

impl<'a> ExcelWriteFillExecutor<'a> {
    /// Creates the executor. (Java `ExcelWriteFillExecutor(WriteContext)`)
    #[must_use]
    pub const fn new(write_context: &'a dyn WriteContext) -> Self {
        Self {
            inner: AbstractExcelWriteExecutor::new(write_context),
            delegate: None,
        }
    }

    /// Creates an executor backed by a real stateful template engine.
    ///
    /// Java constructs the engine directly from `WriteContext`; Rust injects
    /// it to avoid a dependency cycle between `easyexcel-writer` and
    /// `easyexcel-template`.
    #[must_use]
    pub fn with_delegate(
        write_context: &'a dyn WriteContext,
        delegate: &'a mut dyn WriteFillExecutor,
    ) -> Self {
        Self {
            inner: AbstractExcelWriteExecutor::new(write_context),
            delegate: Some(delegate),
        }
    }

    /// Returns the inner `WriteContext`. (Java `getWriteContext()` step)
    #[must_use]
    pub const fn write_context(&self) -> &dyn WriteContext {
        self.inner.write_context
    }

    /// Fills scalar or collection data on one selected worksheet.
    ///
    /// Mirrors Java `fill(Object, FillConfig)`. Template analysis, repeated
    /// collection cursors, row shifting, horizontal/vertical direction and
    /// style inheritance execute in the injected stateful engine.
    ///
    /// # Errors
    ///
    /// Returns a configuration error when no engine is attached, or the real
    /// template conversion/package error from the engine.
    pub fn fill(
        &mut self,
        data: &dyn Any,
        fill_config: WriteFillConfig,
        sheet: WriteFillSheet,
    ) -> Result<()> {
        self.delegate
            .as_deref_mut()
            .ok_or_else(Self::missing_delegate_error)?
            .fill(data, fill_config, sheet)
    }

    /// Persists the accumulated template session.
    ///
    /// Mirrors completion through Java `WriteContext.finish(onException)`.
    ///
    /// # Errors
    ///
    /// Returns a configuration or backend finalization error.
    pub fn finish(&mut self, on_exception: bool) -> Result<()> {
        self.delegate
            .as_deref_mut()
            .ok_or_else(Self::missing_delegate_error)?
            .finish(on_exception)
    }

    fn missing_delegate_error() -> ExcelError {
        ExcelError::Unsupported(
            "template fill executor is not wired to a stateful template engine".to_owned(),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use easyexcel_core::{WriteContextImpl, WriteDirection};

    #[derive(Default)]
    struct ProbeFillExecutor {
        fills: Vec<(WriteFillConfig, WriteFillSheet)>,
        finished: Vec<bool>,
    }

    impl WriteFillExecutor for ProbeFillExecutor {
        fn fill(
            &mut self,
            _data: &dyn Any,
            fill_config: WriteFillConfig,
            sheet: WriteFillSheet,
        ) -> Result<()> {
            self.fills.push((fill_config, sheet));
            Ok(())
        }

        fn finish(&mut self, on_exception: bool) -> Result<()> {
            self.finished.push(on_exception);
            Ok(())
        }
    }

    #[test]
    fn executor_delegates_fill_state_and_finish_to_real_engine() {
        let context = WriteContextImpl::new("output.xlsx");
        let mut probe = ProbeFillExecutor::default();
        {
            let mut executor = ExcelWriteFillExecutor::with_delegate(&context, &mut probe);
            assert_eq!(
                executor.write_context().current_write_holder().path(),
                std::path::Path::new("output.xlsx")
            );
            executor
                .fill(
                    &"payload",
                    WriteFillConfig {
                        force_new_row: true,
                        direction: Some(WriteDirection::Horizontal),
                        auto_style: false,
                    },
                    WriteFillSheet {
                        sheet_name: "Data".to_owned(),
                        sheet_index: Some(2),
                    },
                )
                .expect("delegate fill");
            executor.finish(true).expect("delegate finish");
        }
        assert_eq!(probe.fills.len(), 1);
        assert_eq!(probe.fills[0].0.direction, Some(WriteDirection::Horizontal));
        assert_eq!(probe.fills[0].1.sheet_index, Some(2));
        assert_eq!(probe.finished, vec![true]);
    }

    #[test]
    fn executor_without_engine_returns_visible_error() {
        let context = WriteContextImpl::new("output.xlsx");
        let mut executor = ExcelWriteFillExecutor::new(&context);
        let error = executor
            .fill(
                &"payload",
                WriteFillConfig::default(),
                WriteFillSheet::default(),
            )
            .expect_err("missing engine");
        assert!(error.to_string().contains("not wired"));
    }
}
