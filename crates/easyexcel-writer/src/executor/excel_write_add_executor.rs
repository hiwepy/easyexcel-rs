//! Mirrors Java `com.alibaba.excel.write.executor.ExcelWriteAddExecutor`.
//!
//! Java source:
//! `easyexcel-core/src/main/java/com/alibaba/excel/write/executor/ExcelWriteAddExecutor.java`
//!
//! Method map (Java → Rust):
//! - `add(Collection<?>)` → [`ExcelWriteAddExecutor::add`] /
//!   [`ExcelWriteAddExecutor::add_to_worksheet`]
//! - `addOneRowOfDataToExcel` → [`ExcelWriteAddExecutor::add_one_row_of_data_to_excel`]
//! - `addBasicTypeToExcel` → [`ExcelWriteAddExecutor::add_basic_type_to_excel`] /
//!   [`ExcelWriteAddExecutor::add_basic_type_to_excel_with_map`]
//! - `doAddBasicTypeToExcel` → [`ExcelWriteAddExecutor::do_add_basic_type_to_excel`]
//! - `addJavaObjectToExcel` → [`ExcelWriteAddExecutor::add_java_object_to_excel`]
//!
//! Heavy lifting stays in `lib.rs` (`write_xlsx` / `append_rows_to_worksheet`);
//! this type only adds the Java-named entry points (只增不减).

use std::collections::BTreeMap;
use std::path::Path;

use easyexcel_core::{
    CellValue, DynamicRow, DynamicValue, ExcelRow, ExcelWriteMetadata, Result, WriteContext,
    WriteHandler,
};
use rust_xlsxwriter::Worksheet;

use crate::executor::abstract_excel_write_executor::AbstractExcelWriteExecutor;
use crate::metadata::collection_row_data::CollectionRowData;
use crate::metadata::map_row_data::MapRowData;
use crate::{WriteOptions, WriteProgress, append_rows_to_worksheet, write_xlsx};

/// Mirrors Java `ExcelWriteAddExecutor extends AbstractExcelWriteExecutor`.
///
/// The Java side holds `add(Collection<?>)`, `addOneRowOfDataToExcel`,
/// `addBasicTypeToExcel`, `doAddBasicTypeToExcel`, and `addJavaObjectToExcel`.
/// Rust keeps the same method names (snake_case) and delegates to the existing
/// `write_xlsx` / `append_rows_to_worksheet` writer path so `lib.rs` behaviour
/// is preserved.
pub struct ExcelWriteAddExecutor<'a> {
    inner: AbstractExcelWriteExecutor<'a>,
}

impl<'a> ExcelWriteAddExecutor<'a> {
    /// Creates the executor. (Java `ExcelWriteAddExecutor(WriteContext)`)
    #[must_use]
    pub const fn new(write_context: &'a dyn WriteContext) -> Self {
        Self {
            inner: AbstractExcelWriteExecutor::new(write_context),
        }
    }

    /// Returns the inner `WriteContext`. (Java `getWriteContext()` step)
    #[must_use]
    pub const fn write_context(&self) -> &dyn WriteContext {
        self.inner.write_context
    }

    /// Adds a collection of rows to the workbook path carried by [`WriteContext`].
    ///
    /// Corresponds to Java
    /// `ExcelWriteAddExecutor.add(Collection<?> data)` in
    /// `ExcelWriteAddExecutor.java`.
    ///
    /// # Logic (aligned with Java)
    /// 1. Empty input is treated as an empty list (Java replaces null/empty
    ///    with `new ArrayList<>()` and still runs the loop).
    /// 2. Persist via the existing `write_xlsx` → `write_sheet_to_workbook` →
    ///    `append_rows_to_worksheet` path (no duplicate write engine).
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, XLSX-format, or I/O error from `write_xlsx`.
    pub fn add<T, I>(&self, options: &WriteOptions, data: I) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        // Java: CollectionUtils.isEmpty(data) → data = new ArrayList<>()
        // Rust: IntoIterator of zero items is already an empty collection.
        let path = self.write_context().current_write_holder().path();
        write_xlsx(path, options, data)
    }

    /// Adds a collection of rows onto an existing worksheet.
    ///
    /// Same Java method as [`Self::add`], but for the stateful writer path that
    /// already owns a `Worksheet` (mirrors the loop body that repeatedly calls
    /// `addOneRowOfDataToExcel` after `getNewRowIndexAndStartDoWrite()`).
    ///
    /// Delegates directly to [`append_rows_to_worksheet`].
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or XLSX-format error.
    #[allow(clippy::too_many_arguments)]
    pub fn add_to_worksheet<T, I>(
        &self,
        worksheet: &mut Worksheet,
        options: &WriteOptions,
        data: I,
        handlers: &mut [Box<dyn WriteHandler>],
        progress: WriteProgress,
        write_head: bool,
        metadata: &ExcelWriteMetadata,
    ) -> Result<WriteProgress>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        // Java add(): for (Object oneRowData : data) addOneRowOfDataToExcel(...)
        append_rows_to_worksheet::<T, I>(
            worksheet, options, data, handlers, progress, write_head, metadata,
        )
    }

    /// Writes a single data row.
    ///
    /// Corresponds to Java private
    /// `addOneRowOfDataToExcel(Object oneRowData, int rowIndex, int relativeRowIndex)`.
    ///
    /// # Logic (aligned with Java)
    /// 1. Null / missing row is a no-op in Java; Rust callers simply omit the call.
    /// 2. Row/cell handler hooks run inside `append_rows_to_worksheet` /
    ///    `write_data_row_with_handlers` (Java `WriteHandlerUtils.beforeRowCreate`
    ///    / `afterRowCreate` / `afterRowDispose`).
    /// 3. Collection / Map / JavaBean branching is expressed by the `T: ExcelRow`
    ///    impl (`DynamicRow` for basic types, derived structs for beans).
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or XLSX-format error.
    #[allow(clippy::too_many_arguments)]
    pub fn add_one_row_of_data_to_excel<T>(
        &self,
        worksheet: &mut Worksheet,
        options: &WriteOptions,
        one_row_data: T,
        row_index: u32,
        relative_row_index: usize,
        handlers: &mut [Box<dyn WriteHandler>],
        metadata: &ExcelWriteMetadata,
    ) -> Result<WriteProgress>
    where
        T: ExcelRow,
    {
        // Java: WorkBookUtil.createRow + branch Collection/Map/JavaBean
        // Rust: single-item append at the requested physical row index.
        append_rows_to_worksheet(
            worksheet,
            options,
            std::iter::once(one_row_data),
            handlers,
            WriteProgress {
                next_row: row_index,
                next_data_index: relative_row_index,
            },
            false,
            metadata,
        )
    }

    /// Writes a no-model / basic-type row from a list of cell values.
    ///
    /// Corresponds to Java private
    /// `addBasicTypeToExcel(RowData oneRowData, Row row, int rowIndex, int relativeRowIndex)`
    /// when `oneRowData` is a `CollectionRowData`.
    ///
    /// # Logic (aligned with Java)
    /// 1. Empty row → return immediately (`oneRowData.isEmpty()`).
    /// 2. Convert list cells into a [`DynamicRow`] and reuse
    ///    [`Self::add_one_row_of_data_to_excel`].
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or XLSX-format error.
    #[allow(clippy::too_many_arguments)]
    pub fn add_basic_type_to_excel(
        &self,
        worksheet: &mut Worksheet,
        options: &WriteOptions,
        one_row_data: &CollectionRowData,
        row_index: u32,
        relative_row_index: usize,
        handlers: &mut [Box<dyn WriteHandler>],
        metadata: &ExcelWriteMetadata,
    ) -> Result<WriteProgress> {
        // Java: if (oneRowData.isEmpty()) return;
        if one_row_data.is_empty() {
            return Ok(WriteProgress {
                next_row: row_index,
                next_data_index: relative_row_index,
            });
        }
        let dynamic = collection_row_to_dynamic(one_row_data);
        self.add_one_row_of_data_to_excel(
            worksheet,
            options,
            dynamic,
            row_index,
            relative_row_index,
            handlers,
            metadata,
        )
    }

    /// Writes a no-model / basic-type row from a column-indexed map.
    ///
    /// Corresponds to the Java `Map` branch of
    /// `addBasicTypeToExcel` / `addOneRowOfDataToExcel` (`MapRowData`).
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or XLSX-format error.
    #[allow(clippy::too_many_arguments)]
    pub fn add_basic_type_to_excel_with_map(
        &self,
        worksheet: &mut Worksheet,
        options: &WriteOptions,
        one_row_data: &MapRowData,
        row_index: u32,
        relative_row_index: usize,
        handlers: &mut [Box<dyn WriteHandler>],
        metadata: &ExcelWriteMetadata,
    ) -> Result<WriteProgress> {
        // Java: if (oneRowData.isEmpty()) return;
        if one_row_data.is_empty() {
            return Ok(WriteProgress {
                next_row: row_index,
                next_data_index: relative_row_index,
            });
        }
        let dynamic = map_row_to_dynamic(one_row_data);
        self.add_one_row_of_data_to_excel(
            worksheet,
            options,
            dynamic,
            row_index,
            relative_row_index,
            handlers,
            metadata,
        )
    }

    /// Writes a single basic-type cell into a new row at the given column.
    ///
    /// Corresponds to Java private
    /// `doAddBasicTypeToExcel(RowData, Head, Row, int, int, int, int)`.
    ///
    /// # Logic (aligned with Java)
    /// 1. Java creates a cell, runs converters, then `converterAndSet`.
    /// 2. Rust packs the value into a sparse [`DynamicRow`] keyed by
    ///    `column_index` and reuses [`Self::add_one_row_of_data_to_excel`]
    ///    (same `append_rows_to_worksheet` engine).
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or XLSX-format error.
    #[allow(clippy::too_many_arguments)]
    pub fn do_add_basic_type_to_excel(
        &self,
        worksheet: &mut Worksheet,
        options: &WriteOptions,
        value: CellValue,
        row_index: u32,
        relative_row_index: usize,
        column_index: usize,
        handlers: &mut [Box<dyn WriteHandler>],
        metadata: &ExcelWriteMetadata,
    ) -> Result<WriteProgress> {
        // Java: WorkBookUtil.createCell(row, columnIndex) + converterAndSet
        let mut cells = BTreeMap::new();
        cells.insert(column_index, DynamicValue::ActualData(value));
        self.add_one_row_of_data_to_excel(
            worksheet,
            options,
            DynamicRow::new(cells),
            row_index,
            relative_row_index,
            handlers,
            metadata,
        )
    }

    /// Writes a typed JavaBean-equivalent row (`ExcelRow` derive / impl).
    ///
    /// Corresponds to Java private
    /// `addJavaObjectToExcel(Object oneRowData, Row row, int rowIndex, int relativeRowIndex)`.
    ///
    /// # Logic (aligned with Java)
    /// 1. Java reflects fields via `BeanMap` + `Head` / `FieldCache`.
    /// 2. Rust uses `T: ExcelRow` (`to_row` / converters) and delegates to
    ///    [`Self::add_one_row_of_data_to_excel`].
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, or XLSX-format error.
    #[allow(clippy::too_many_arguments)]
    pub fn add_java_object_to_excel<T>(
        &self,
        worksheet: &mut Worksheet,
        options: &WriteOptions,
        one_row_data: T,
        row_index: u32,
        relative_row_index: usize,
        handlers: &mut [Box<dyn WriteHandler>],
        metadata: &ExcelWriteMetadata,
    ) -> Result<WriteProgress>
    where
        T: ExcelRow,
    {
        self.add_one_row_of_data_to_excel(
            worksheet,
            options,
            one_row_data,
            row_index,
            relative_row_index,
            handlers,
            metadata,
        )
    }

    /// Convenience used by tests / callers that already have an explicit path
    /// rather than a [`WriteContext`] holder path.
    ///
    /// Same write engine as [`Self::add`]; does not replace it.
    ///
    /// # Errors
    ///
    /// Returns a conversion, handler, XLSX-format, or I/O error from `write_xlsx`.
    pub fn add_to_path<T, I>(&self, path: &Path, options: &WriteOptions, data: I) -> Result<()>
    where
        T: ExcelRow,
        I: IntoIterator<Item = T>,
    {
        write_xlsx(path, options, data)
    }
}

/// Converts Java `CollectionRowData` into a writeable [`DynamicRow`].
fn collection_row_to_dynamic(row: &CollectionRowData) -> DynamicRow {
    let cells = row
        .values()
        .iter()
        .enumerate()
        .map(|(index, value)| (index, DynamicValue::ActualData(value.clone())))
        .collect();
    DynamicRow::new(cells)
}

/// Converts Java `MapRowData` into a writeable [`DynamicRow`].
fn map_row_to_dynamic(row: &MapRowData) -> DynamicRow {
    // Java `MapRowData.size()` returns the map entry count, then
    // `ExcelWriteAddExecutor` calls `get(dataIndex)` for 0..size. It does not
    // iterate arbitrary integer keys as physical column indexes.
    let cells = (0..row.values().len())
        .map(|index| {
            (
                index,
                DynamicValue::ActualData(
                    row.values()
                        .get(&index)
                        .cloned()
                        .unwrap_or(CellValue::Empty),
                ),
            )
        })
        .collect();
    DynamicRow::new(cells)
}

#[cfg(test)]
mod tests {
    use std::path::{Path, PathBuf};

    use calamine::{Data, Reader, Xlsx, open_workbook};
    use easyexcel_core::{WriteContext, WriteContextHolder};
    use rust_xlsxwriter::Workbook;
    use tempfile::tempdir;

    use super::*;

    /// Minimal `WriteContext` whose holder path is the output XLSX file.
    struct TestWriteContext {
        path: PathBuf,
    }

    impl WriteContext for TestWriteContext {
        fn current_write_holder(&self) -> &dyn WriteContextHolder {
            self
        }
    }

    impl WriteContextHolder for TestWriteContext {
        fn path(&self) -> &Path {
            &self.path
        }

        fn holder_type(&self) -> easyexcel_core::Holder {
            easyexcel_core::Holder::Workbook
        }

        fn excel_write_head_property(&self) -> &easyexcel_core::ExcelWriteHeadProperty {
            static PROPERTY: std::sync::OnceLock<easyexcel_core::ExcelWriteHeadProperty> =
                std::sync::OnceLock::new();
            PROPERTY.get_or_init(easyexcel_core::ExcelWriteHeadProperty::new)
        }

        fn converter_map(&self) -> &easyexcel_core::ConverterRegistry {
            static REGISTRY: std::sync::OnceLock<easyexcel_core::ConverterRegistry> =
                std::sync::OnceLock::new();
            REGISTRY.get_or_init(easyexcel_core::ConverterRegistry::default)
        }

        fn need_head(&self) -> bool {
            true
        }

        fn automatic_merge_head(&self) -> bool {
            true
        }

        fn relative_head_row_index(&self) -> i32 {
            0
        }

        fn order_by_include_column(&self) -> bool {
            false
        }

        fn include_column_indexes(&self) -> Option<&[usize]> {
            None
        }

        fn include_column_field_names(&self) -> Option<&[String]> {
            None
        }

        fn exclude_column_indexes(&self) -> &[usize] {
            &[]
        }

        fn exclude_column_field_names(&self) -> &[String] {
            &[]
        }
    }

    /// Proves Java `add(Collection)` is callable and writes a calamine-readable file.
    #[test]
    fn add_writes_readable_xlsx() {
        let directory = tempdir().expect("tempdir");
        let path = directory.path().join("add-executor.xlsx");
        let context = TestWriteContext { path: path.clone() };
        let executor = ExcelWriteAddExecutor::new(&context);

        let mut row = BTreeMap::new();
        row.insert(0, DynamicValue::String("alice".to_owned()));
        row.insert(1, DynamicValue::ActualData(CellValue::Int(18)));
        let data = vec![DynamicRow::new(row)];

        executor
            .add(
                &WriteOptions {
                    need_head: false,
                    sheet_name: "Sheet1".to_owned(),
                    ..WriteOptions::default()
                },
                data,
            )
            .expect("add should succeed");

        let mut workbook: Xlsx<_> = open_workbook(&path).expect("open written xlsx");
        let range = workbook.worksheet_range("Sheet1").expect("sheet");
        assert_eq!(range.get((0, 0)), Some(&Data::String("alice".to_owned())));
        assert_eq!(range.get((0, 1)), Some(&Data::Float(18.0)));
    }

    /// Proves `add_one_row_of_data_to_excel` + basic-type helpers write cells.
    #[test]
    fn add_one_row_and_basic_type_write_cells() {
        let directory = tempdir().expect("tempdir");
        let path = directory.path().join("one-row.xlsx");
        let context = TestWriteContext { path: path.clone() };
        let executor = ExcelWriteAddExecutor::new(&context);

        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("Data").expect("sheet name");
        let options = WriteOptions {
            need_head: false,
            sheet_name: "Data".to_owned(),
            ..WriteOptions::default()
        };
        let metadata = ExcelWriteMetadata::new();

        // Java addOneRowOfDataToExcel / addBasicTypeToExcel (CollectionRowData)
        let collection = CollectionRowData::new(vec![
            CellValue::String("bob".to_owned()),
            CellValue::Int(21),
        ]);
        let progress = executor
            .add_basic_type_to_excel(worksheet, &options, &collection, 0, 0, &mut [], &metadata)
            .expect("collection row");
        assert_eq!(progress.next_row, 1);

        // Java MapRowData uses map.size() + get(0..size), not sparse key
        // iteration. Key 2 is therefore outside the two-entry row.
        let mut map = BTreeMap::new();
        map.insert(0, CellValue::String("carol".to_owned()));
        map.insert(2, CellValue::Int(30));
        let map_row = MapRowData::new(map);
        let progress = executor
            .add_basic_type_to_excel_with_map(
                worksheet,
                &options,
                &map_row,
                1,
                1,
                &mut [],
                &metadata,
            )
            .expect("map row");
        assert_eq!(progress.next_row, 2);

        // Java doAddBasicTypeToExcel — single cell at column 1
        let progress = executor
            .do_add_basic_type_to_excel(
                worksheet,
                &options,
                CellValue::String("solo".to_owned()),
                2,
                2,
                1,
                &mut [],
                &metadata,
            )
            .expect("single cell");
        assert_eq!(progress.next_row, 3);

        crate::save_workbook(&mut workbook, &path, None).expect("save");

        let mut book: Xlsx<_> = open_workbook(&path).expect("open");
        let range = book.worksheet_range("Data").expect("sheet");
        assert_eq!(range.get((0, 0)), Some(&Data::String("bob".to_owned())));
        assert_eq!(range.get((0, 1)), Some(&Data::Float(21.0)));
        assert_eq!(range.get((1, 0)), Some(&Data::String("carol".to_owned())));
        assert_eq!(range.get((1, 1)), Some(&Data::Empty));
        assert_eq!(range.get((1, 2)), None);
        assert_eq!(range.get((2, 1)), Some(&Data::String("solo".to_owned())));
    }
}
