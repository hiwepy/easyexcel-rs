//! Mirrors the registry half of Java `com.alibaba.excel.converters.ConverterKeyBuild`
//! and `com.alibaba.excel.converters.DefaultConverterLoader`.

use std::any::{Any, TypeId, type_name};
use std::fmt;
use std::sync::Arc;

use crate::WriteCellData;
use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::converter::Converter;
use crate::converter::converter_key_build::ConverterKey;
use crate::converter::nullable_object_converter::NullableObjectConverter;
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
    fn write_target_type(&self) -> Option<CellDataType>;
    fn accepts_null(&self) -> bool;
    fn convert_to_rust_data(
        &self,
        context: &ReadConverterContext<'_>,
    ) -> Result<Box<dyn Any>, ExcelError>;
    fn convert_to_excel_data(
        &self,
        value: &dyn Any,
        column: &ExcelColumn,
        context: &ConvertContext,
    ) -> Result<WriteCellData, ExcelError>;
}

/// Type-tagged carrier for `Converter<T>`.
///
/// Mirrors `TypedConverter` from the Java side. The marker phantom
/// parameter is the Rust equivalent of `Converter<T>.supportJavaTypeKey()`.
pub(crate) struct TypedConverter<T, C> {
    pub(crate) converter: C,
    pub(crate) write_target_type: Option<CellDataType>,
    pub(crate) accepts_null: bool,
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

    fn write_target_type(&self) -> Option<CellDataType> {
        self.write_target_type
    }

    fn accepts_null(&self) -> bool {
        self.accepts_null
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
    ) -> Result<WriteCellData, ExcelError> {
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
    requested_write_type: Option<CellDataType>,
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
            write_target_type: None,
            accepts_null: false,
            marker: std::marker::PhantomData,
        }));
    }

    /// Registers a converter under Java's `(class, targetCellDataType)` write key.
    pub fn register_for_write_type<T, C>(&mut self, target: CellDataType, converter: C)
    where
        T: 'static,
        C: Converter<T> + Send + Sync + 'static,
    {
        self.converters.push(Arc::new(TypedConverter::<T, C> {
            converter,
            write_target_type: Some(target),
            accepts_null: false,
            marker: std::marker::PhantomData,
        }));
    }

    /// Registers Java's `NullableObjectConverter<T>` under the normal read/write key.
    ///
    /// Unlike an ordinary converter, this converter is invoked for an empty
    /// source cell and may be selected for an absent `Option<T>` write.
    pub fn register_nullable<T, C>(&mut self, converter: C)
    where
        T: 'static,
        C: NullableObjectConverter<T> + Send + Sync + 'static,
    {
        self.converters.push(Arc::new(TypedConverter::<T, C> {
            converter,
            write_target_type: None,
            accepts_null: true,
            marker: std::marker::PhantomData,
        }));
    }

    /// Registers a nullable converter under a target Excel cell type.
    pub fn register_nullable_for_write_type<T, C>(&mut self, target: CellDataType, converter: C)
    where
        T: 'static,
        C: NullableObjectConverter<T> + Send + Sync + 'static,
    {
        self.converters.push(Arc::new(TypedConverter::<T, C> {
            converter,
            write_target_type: Some(target),
            accepts_null: true,
            marker: std::marker::PhantomData,
        }));
    }

    /// Returns a registry where `overrides` take precedence over this registry.
    /// Mirrors Java sheet-level > workbook-level converter chain.
    #[must_use]
    pub fn merged_with(&self, overrides: &Self) -> Self {
        let mut converters = self.converters.clone();
        converters.extend(overrides.converters.iter().cloned());
        Self {
            converters,
            requested_write_type: overrides.requested_write_type.or(self.requested_write_type),
        }
    }

    /// Returns a clone selecting Java's target cell type for this write pass.
    #[must_use]
    pub fn with_write_target(&self, target: Option<CellDataType>) -> Self {
        let mut registry = self.clone();
        registry.requested_write_type = target;
        registry
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
        let requested_key = ConverterKey::of::<T>(Some(data_type));
        let Some(converter) = self.converters.iter().rev().find(|converter| {
            ConverterKey::new(
                converter.target_type_id(),
                Some(converter.support_excel_type()),
            ) == requested_key
        }) else {
            return Ok(None);
        };
        if data_type == CellDataType::Empty && !converter.accepts_null() {
            return Ok(None);
        }
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
    ) -> Result<Option<WriteCellData>, ExcelError>
    where
        T: 'static,
    {
        self.convert_to_excel_data_with_null_state(value, column, context, false)
    }

    /// Converts a Rust value while preserving Java's nullable-converter gate.
    ///
    /// Derive-generated `Option<T>` fields pass `true` for `value_is_null`.
    /// An ordinary converter is skipped and the caller writes an empty cell;
    /// a converter registered through [`Self::register_nullable`] is invoked.
    pub fn convert_to_excel_data_with_null_state<T>(
        &self,
        value: &T,
        column: &ExcelColumn,
        context: &ConvertContext,
        value_is_null: bool,
    ) -> Result<Option<WriteCellData>, ExcelError>
    where
        T: 'static,
    {
        let requested_key = ConverterKey::of::<T>(self.requested_write_type);
        let Some(converter) = self.converters.iter().rev().find(|converter| {
            let exact_key =
                ConverterKey::new(converter.target_type_id(), converter.write_target_type());
            exact_key == requested_key
                || (converter.target_type_id() == requested_key.rust_type()
                    && converter.write_target_type().is_none())
        }) else {
            return Ok(None);
        };
        if value_is_null && !converter.accepts_null() {
            return Ok(None);
        }
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
            .entries(self.converters.iter().map(|converter| {
                (
                    converter.target_type_name(),
                    converter.support_excel_type(),
                    converter.write_target_type(),
                )
            }))
            .finish()
    }
}

impl PartialEq for ConverterRegistry {
    fn eq(&self, other: &Self) -> bool {
        self.requested_write_type == other.requested_write_type
            && self.converters.len() == other.converters.len()
            && self
                .converters
                .iter()
                .zip(&other.converters)
                .all(|(left, right)| {
                    left.target_type_id() == right.target_type_id()
                        && left.support_excel_type() == right.support_excel_type()
                        && left.write_target_type() == right.write_target_type()
                        && left.accepts_null() == right.accepts_null()
                })
    }
}

impl Eq for ConverterRegistry {}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::converter::nullable_object_converter::NullableObjectConverter;
    use crate::{CellValue, ConvertContext, ExcelColumn, ReadConverterContext};

    #[derive(Clone, Copy)]
    struct EmptyStringConverter;

    impl Converter<String> for EmptyStringConverter {
        fn support_excel_type(&self) -> CellDataType {
            CellDataType::Empty
        }

        fn convert_to_rust_data(
            &self,
            _context: &ReadConverterContext<'_>,
        ) -> Result<String, ExcelError> {
            Ok("converted-empty".to_owned())
        }
    }

    #[derive(Clone, Copy)]
    struct NullableEmptyStringConverter;

    impl Converter<String> for NullableEmptyStringConverter {
        fn support_excel_type(&self) -> CellDataType {
            CellDataType::Empty
        }

        fn convert_to_rust_data(
            &self,
            _context: &ReadConverterContext<'_>,
        ) -> Result<String, ExcelError> {
            Ok("converted-empty".to_owned())
        }
    }

    impl NullableObjectConverter<String> for NullableEmptyStringConverter {}

    #[derive(Clone, Copy)]
    struct OptionWriteConverter;

    impl Converter<Option<String>> for OptionWriteConverter {
        fn convert_to_excel_data(
            &self,
            context: &WriteConverterContext<'_, Option<String>>,
        ) -> Result<WriteCellData, ExcelError> {
            Ok(WriteCellData::new(CellValue::String(
                context
                    .value()
                    .as_deref()
                    .unwrap_or("ordinary-null")
                    .to_owned(),
            )))
        }
    }

    #[derive(Clone, Copy)]
    struct NullableOptionWriteConverter;

    impl Converter<Option<String>> for NullableOptionWriteConverter {
        fn convert_to_excel_data(
            &self,
            context: &WriteConverterContext<'_, Option<String>>,
        ) -> Result<WriteCellData, ExcelError> {
            Ok(WriteCellData::new(CellValue::String(
                context
                    .value()
                    .as_deref()
                    .unwrap_or("nullable-null")
                    .to_owned(),
            )))
        }
    }

    impl NullableObjectConverter<Option<String>> for NullableOptionWriteConverter {}

    fn location() -> ConvertContext {
        ConvertContext {
            sheet_name: "Data".to_owned(),
            row_index: 1,
            column_index: Some(0),
            field: "value",
            format: None,
            use_1904_windowing: false,
        }
    }

    #[test]
    fn empty_read_invokes_only_nullable_object_converter() {
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        let context = location();
        let read_context = ReadConverterContext::new(None, &column, &context);

        let mut ordinary = ConverterRegistry::default();
        ordinary.register::<String, _>(EmptyStringConverter);
        assert_eq!(
            ordinary
                .convert_to_rust_data::<String>(&read_context)
                .expect("ordinary converter dispatch"),
            None
        );

        let mut nullable = ConverterRegistry::default();
        nullable.register_nullable::<String, _>(NullableEmptyStringConverter);
        assert_eq!(
            nullable
                .convert_to_rust_data::<String>(&read_context)
                .expect("nullable converter dispatch"),
            Some("converted-empty".to_owned())
        );
    }

    #[test]
    fn absent_option_write_invokes_only_nullable_object_converter() {
        let column = ExcelColumn::new("value", "Value", Some(0), 0, None);
        let context = location();
        let value: Option<String> = None;

        let mut ordinary = ConverterRegistry::default();
        ordinary.register::<Option<String>, _>(OptionWriteConverter);
        assert!(
            ordinary
                .convert_to_excel_data_with_null_state(&value, &column, &context, true)
                .expect("ordinary converter dispatch")
                .is_none()
        );

        let mut nullable = ConverterRegistry::default();
        nullable.register_nullable::<Option<String>, _>(NullableOptionWriteConverter);
        let converted = nullable
            .convert_to_excel_data_with_null_state(&value, &column, &context, true)
            .expect("nullable converter dispatch")
            .expect("nullable converter selected");
        assert_eq!(
            converted.value(),
            &CellValue::String("nullable-null".to_owned())
        );
    }
}
