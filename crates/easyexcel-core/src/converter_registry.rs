//! Mirrors the registry half of Java `com.alibaba.excel.converters.ConverterKeyBuild`
//! and `com.alibaba.excel.converters.DefaultConverterLoader`.

use std::any::{Any, TypeId, type_name};
use std::fmt;
use std::sync::Arc;

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::converter::Converter;
use crate::enum_cell_data_type::CellDataType;
use crate::excel_column::ExcelColumn;
use crate::excel_error::ExcelError;
use crate::read_converter_context::ReadConverterContext;
use crate::write_converter_context::WriteConverterContext;

/// Trait-object erase of `Converter<T>` keyed by `TypeId`.
///
/// Mirrors the role of `ConverterKeyBuild.ConverterKey` plus the dispatch
/// through `ConverterKeyBuild.buildKey(Class, CellDataTypeEnum)`. Rust uses
/// `TypeId` because `TypeId` is the type-safe `Class` equivalent.
pub(crate) trait ErasedConverter: Send + Sync {
    fn target_type_id(&self) -> TypeId;
    fn target_type_name(&self) -> &'static str;
    fn support_excel_type(&self) -> CellDataType;
    fn convert_to_rust_data(
        &self,
        context: &ReadConverterContext<'_>,
    ) -> Result<Box<dyn Any>, ExcelError>;
    fn convert_to_excel_data(
        &self,
        value: &dyn Any,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<CellValue, ExcelError>;
}

/// Type-tagged carrier for `Converter<T>`.
///
/// Mirrors `TypedConverter` from the Java side. The marker phantom
/// parameter is the Rust equivalent of `Converter<T>.supportJavaTypeKey()`.
pub(crate) struct TypedConverter<T, C> {
    pub(crate) converter: C,
    pub(crate) marker: std::marker::PhantomData<fn() -> T>,
}

impl<T, C> ErasedConverter for TypedConverter<T, C>
where
    T: 'static,
    C: Converter<T> + Send + Sync,
{
    fn target_type_id(&self) -> TypeId {
        TypeId::of::<T>()
    }

    fn target_type_name(&self) -> &'static str {
        type_name::<T>()
    }

    fn support_excel_type(&self) -> CellDataType {
        self.converter.support_excel_type()
    }

    fn convert_to_rust_data(
        &self,
        context: &ReadConverterContext<'_>,
    ) -> Result<Box<dyn Any>, ExcelError> {
        self.converter
            .convert_to_rust_data(context)
            .map(|value| Box::new(value) as Box<dyn Any>)
    }

    fn convert_to_excel_data(
        &self,
        value: &dyn Any,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<CellValue, ExcelError> {
        let value = value.downcast_ref::<T>().ok_or_else(|| {
            ExcelError::Format(format!(
                "registered converter expected Rust type {}",
                type_name::<T>()
            ))
        })?;
        self.converter
            .convert_to_excel_data(&WriteConverterContext::new(value, column, context))
    }
}

/// Runtime converter registry populated by Java-style `registerConverter` builders.
///
/// Mirrors the union of `DefaultConverterLoader` and `AbstractHolder.converterMap`.
/// Registrations are searched from newest to oldest. Read selection uses the
/// pair `(Rust target type, Excel cell type)` while write selection uses only
/// the Rust type, matching Java `EasyExcel`'s holder initialization rules.
#[derive(Clone, Default)]
pub struct ConverterRegistry {
    pub(crate) converters: Vec<Arc<dyn ErasedConverter>>,
}

impl ConverterRegistry {
    /// Registers a converter for `T`, overriding an earlier converter with the same key.
    /// (Java `DefaultConverterLoader.putWriteConverter`)
    pub fn register<T, C>(&mut self, converter: C)
    where
        T: 'static,
        C: Converter<T> + Send + Sync + 'static,
    {
        self.converters.push(Arc::new(TypedConverter::<T, C> {
            converter,
            marker: std::marker::PhantomData,
        }));
    }

    /// Returns a registry where `overrides` take precedence over this registry.
    /// Mirrors Java sheet-level > workbook-level converter chain.
    #[must_use]
    pub fn merged_with(&self, overrides: &Self) -> Self {
        let mut converters = self.converters.clone();
        converters.extend(overrides.converters.iter().cloned());
        Self { converters }
    }

    /// Converts a cell through the newest matching global converter.
    ///
    /// `None` means no global converter matched and the caller should use its
    /// built-in conversion implementation. Mirrors Java
    /// `AbstractHolder.converterMap` + read dispatch.
    ///
    /// # Errors
    ///
    /// Returns the registered converter's error or a type-contract error.
    pub fn convert_to_rust_data<T>(
        &self,
        context: &ReadConverterContext<'_>,
    ) -> Result<Option<T>, ExcelError>
    where
        T: 'static,
    {
        let data_type = context
            .cell()
            .map_or(CellDataType::Empty, CellValue::data_type);
        let Some(converter) = self.converters.iter().rev().find(|converter| {
            converter.target_type_id() == TypeId::of::<T>()
                && converter.support_excel_type() == data_type
        }) else {
            return Ok(None);
        };
        converter
            .convert_to_rust_data(context)?
            .downcast::<T>()
            .map(|value| Some(*value))
            .map_err(|_| {
                ExcelError::Format(format!(
                    "registered converter returned a value other than {}",
                    type_name::<T>()
                ))
            })
    }

    /// Converts a Rust value through the newest matching global converter. Mirrors
    /// `AbstractHolder.converterMap` + write dispatch.
    ///
    /// # Errors
    ///
    /// Returns the registered converter's conversion error.
    pub fn convert_to_excel_data<T>(
        &self,
        value: &T,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<Option<CellValue>, ExcelError>
    where
        T: 'static,
    {
        let Some(converter) = self
            .converters
            .iter()
            .rev()
            .find(|converter| converter.target_type_id() == TypeId::of::<T>())
        else {
            return Ok(None);
        };
        converter
            .convert_to_excel_data(value, column, context)
            .map(Some)
    }

    /// Returns whether no custom converter has been registered.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.converters.is_empty()
    }
}

impl fmt::Debug for ConverterRegistry {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_list()
            .entries(
                self.converters.iter().map(|converter| {
                    (converter.target_type_name(), converter.support_excel_type())
                }),
            )
            .finish()
    }
}

impl PartialEq for ConverterRegistry {
    fn eq(&self, other: &Self) -> bool {
        self.converters.len() == other.converters.len()
            && self
                .converters
                .iter()
                .zip(&other.converters)
                .all(|(left, right)| {
                    left.target_type_id() == right.target_type_id()
                        && left.support_excel_type() == right.support_excel_type()
                })
    }
}

impl Eq for ConverterRegistry {}
