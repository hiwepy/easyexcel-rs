# easyexcel (Java) vs easyexcel-rust — Object/Method Matrix

> Generated 2026-07-20. Each row is one method declaration. Status legend:
> `[OK]` = same name + same semantics implemented in Rust.
> `[OK-RENAMED]` = Java camelCase migrated to Rust snake_case (idiomatic Rust).
> `[OK-ADAPTED]` = signature adapted to Rust idioms (lifetimes / generics / Result).
> `[GAP]` = not yet implemented in Rust (will be addressed in later phases).
> `[HANDLE]` = design choice made (Rust cannot match Java shape exactly).

---

## A. com.alibaba.excel.EasyExcelFactory (Java) ↔ easyexcel::EasyExcel (Rust)

### A.1 Factory / facade methods

| # | Java method | Rust method | Status |
|---|-------------|-------------|--------|
| 1 | `write()` | `EasyExcel::writer()` | [OK-RENAMED] |
| 2 | `write(File)` | `EasyExcel::writer_to_path(path)` | [OK-ADAPTED] |
| 3 | `write(File, Class)` | `EasyExcel::write::<T>(path)` / generic `T` | [OK-ADAPTED] |
| 4 | `write(String)` | `EasyExcel::writer_to_path(path)` | [OK-ADAPTED] |
| 5 | `write(String, Class)` | `EasyExcel::write::<T>(path)` / generic `T` | [OK-ADAPTED] |
| 6 | `write(OutputStream)` | `EasyExcel::writer_to_output_stream(output)` | [OK-ADAPTED] |
| 7 | `write(OutputStream, Class)` | `EasyExcel::write::<T>(logical_path).to_output_stream(output)` | [OK-ADAPTED] |
| 8 | `writerSheet()` | `EasyExcel::writer_sheet_builder()` | [OK-RENAMED] |
| 9 | `writerSheet(Integer)` | `EasyExcel::writer_sheet_builder_index(index)` | [OK-RENAMED] |
| 10 | `writerSheet(String)` | `EasyExcel::writer_sheet_builder_name(name)` | [OK-RENAMED] |
| 11 | `writerSheet(Integer, String)` | `EasyExcel::writer_sheet_builder_with(index, name)` | [OK-RENAMED] |
| 12 | `writerTable()` | `EasyExcel::writer_table_builder_default()` | [OK-RENAMED] |
| 13 | `writerTable(Integer)` | `EasyExcel::writer_table_builder(index)` | [OK-RENAMED] |
| 14 | `read()` | `EasyExcel::reader()` | [OK-RENAMED] |
| 15 | `read(File)` | `EasyExcel::reader_from_path(path)` | [OK-ADAPTED] |
| 16 | `read(File, ReadListener)` | `reader_from_path(path).register_read_listener(listener)` | [OK-ADAPTED] |
| 17 | `read(File, Class, ReadListener)` | `EasyExcel::read::<T, L>(path, listener)` | [OK-ADAPTED] |
| 18 | `read(String)` | `EasyExcel::reader_from_path(path)` | [OK-ADAPTED] |
| 19 | `read(String, ReadListener)` | `reader_from_path(path).register_read_listener(listener)` | [OK-ADAPTED] |
| 20 | `read(String, Class, ReadListener)` | `EasyExcel::read::<T, L>(path, listener)` | [OK-ADAPTED] |
| 21 | `read(InputStream)` | `EasyExcel::reader_from_input_stream(reader)` | [OK-ADAPTED] |
| 22 | `read(InputStream, ReadListener)` | `reader_from_input_stream(reader)?.register_read_listener(listener)` | [OK-ADAPTED] |
| 23 | `read(InputStream, Class, ReadListener)` | previous builder + generic `T: ExcelRow` | [OK-ADAPTED] |
| 24 | `readSheet()` | `EasyExcel::read_sheet()` | [OK-RENAMED] |
| 25 | `readSheet(Integer)` / `readSheet(String)` | `read_sheet_index` / `read_sheet_name` | [OK-RENAMED] |
| 26 | `readSheet(Integer, String)` | `EasyExcel::read_sheet_with(index, name)` | [OK-RENAMED] |

Java 4.0.3 当前基线没有 `Writer` 或 `readTable` factory overload；旧矩阵中的这些行已移除。

---

## B. com.alibaba.excel.ExcelReader ↔ easyexcel_reader::ExcelReader

| # | Java method                              | Rust method                                  | Status        |
|---|------------------------------------------|----------------------------------------------|---------------|
| 1 | `ExcelReader(ReadWorkbook)` constructor  | `ExcelReader::new(path, options, listener)`  | [OK-ADAPTED]  |
| 2 | `read()` (@Deprecated)                   | `ExcelReader::read_deprecated()`             | [OK-RENAMED]  |
| 3 | `readAll()`                              | `ExcelReader::read_all()`                    | [OK-RENAMED]  |
| 4 | `read(ReadSheet...)`                     | `ExcelReader::read(&[ReadSheet])`            | [OK-ADAPTED]  |
| 5 | `read(List<ReadSheet>)`                  | `ExcelReader::read(&[ReadSheet])`            | [OK-ADAPTED]  |
| 6 | `analysisContext()`                      | `ExcelReader::analysis_context()`             | [OK-RENAMED]  |
| 7 | `getAnalysisContext()` (@Deprecated)     | `ExcelReader::get_analysis_context()`         | [OK-RENAMED]  |
| 8 | `excelExecutor()`                        | `ExcelReader::excel_executor()`              | [OK-RENAMED]  |
| 9 | `finish()`                               | `ExcelReader::finish()`                      | [OK]          |
| 10| `close()`                                | `ExcelReader::close()`                       | [OK]          |
| 11| `finalize()`                             | `Drop` impl                                  | [OK-ADAPTED]  |

---

## C. com.alibaba.excel.ExcelWriter ↔ easyexcel_writer::ExcelWriter

| # | Java method                                              | Rust method                                   | Status        |
|---|----------------------------------------------------------|-----------------------------------------------|---------------|
| 1 | `ExcelWriter(WriteWorkbook)` constructor                 | `ExcelWriter::new(write_workbook)`            | [OK-RENAMED]  |
| 2 | `write(Collection<?>, WriteSheet)`                       | `ExcelWriter::write<I,T>(rows, &sheet)`       | [OK-ADAPTED]  |
| 3 | `write(Supplier<Collection<?>>, WriteSheet)`             | `ExcelWriter::write_supplier(...)`            | [HANDLE]      |
| 4 | `write(Collection<?>, WriteSheet, WriteTable)`           | `ExcelWriter::write_with_table(...)`          | [GAP] Phase 4 |
| 5 | `fill(Object, FillConfig, WriteSheet)`                   | `ExcelWriter::fill(...)` / template writer    | [OK-ADAPTED]  |
| 6 | `fill(WriteSheet)`                                       | (delegated to `EasyExcel::fill_template`)     | [HANDLE]      |
| 7 | `writeContext()`                                         | `ExcelWriter::write_context()`                | [OK-RENAMED]  |
| 8 | `finish()`                                               | `ExcelWriter::finish()`                       | [OK]          |
| 9 | `close()`                                                | `ExcelWriter::close()` (auto)                 | [OK]          |
| 10| `finalize()`                                             | `Drop` impl                                   | [OK-ADAPTED]  |

---

## D. com.alibaba.excel.analysis.ExcelAnalyser ↔ easyexcel_reader::analysis::ExcelAnalyser

| # | Java method                                              | Rust trait method                              | Status        |
|---|----------------------------------------------------------|------------------------------------------------|---------------|
| 1 | `analysis(List<ReadSheet>, Boolean)`                     | `ExcelAnalyser::analysis::<T,L>(&mut listener)` + `ReadOptions` sheet selection | [OK-ADAPTED] |
| 2 | `analysisContext()`                                      | `ExcelAnalyser::analysis_context()`            | [OK-RENAMED]  |
| 3 | `excelExecutor()`                                        | `ExcelAnalyser::excel_executor()`              | [OK-RENAMED]  |
| 4 | `finish()`                                               | `ExcelAnalyser::finish()`                      | [OK]          |

### D.1 ExcelAnalyserImpl (concrete)

| # | Java method                          | Rust method                                  | Status        |
|---|--------------------------------------|----------------------------------------------|---------------|
| 1 | `ExcelAnalyserImpl(ReadWorkbook)`    | `ExcelAnalyserImpl::new()` + `from_path()`   | [OK-ADAPTED]  |
| 2 | `analysis(...)`                      | `ExcelAnalyserImpl::analysis(...)`           | [OK-RENAMED]  |
| 3 | `analysisContext()`                  | `ExcelAnalyserImpl::analysis_context()`       | [OK-RENAMED]  |
| 4 | `excelExecutor()`                    | `ExcelAnalyserImpl::excel_executor()`         | [OK-RENAMED]  |
| 5 | `finish()`                           | `ExcelAnalyserImpl::finish()`                 | [OK-ADAPTED]  |
| 6 | `choiceExcelExecutor()`              | `ExcelAnalyserImpl::choice_excel_executor()` | [OK-RENAMED]  |
| 7 | `removeThreadLocalCache()`           | `finish()` 清理 formatter 与磁盘缓存的真实 TLS 句柄；日期/字段元数据无运行时缓存 | [OK-ADAPTED] |
| 8 | `clearEncrypt03()`                   | 密码只属于单个 `ReadOptions`，从未写入 POI 式全局/TLS 状态 | [OK-ADAPTED]  |

### D.2 ExcelReadExecutor / CsvExcelReadExecutor

| Java method | Rust method | Status |
|-------------|-------------|--------|
| `ExcelReadExecutor.sheetList()` | `ExcelReadExecutor::sheet_list()` | [OK-RENAMED] |
| `ExcelReadExecutor.execute()` | `ExcelReadExecutor::execute::<T,L>(options, listener)` | [OK-ADAPTED] |
| `CsvExcelReadExecutor(CsvReadContext)` | `CsvExcelReadExecutor::from_path(path)` + `ReadOptions` | [OK-ADAPTED] |
| `CsvExcelReadExecutor.sheetList()` | `sheet_list()` | [OK-RENAMED] |
| `CsvExcelReadExecutor.execute()` | typed `execute()` using the real CSV parser | [OK-ADAPTED] |

---

## E. com.alibaba.excel.write.ExcelBuilder ↔ easyexcel_writer::excel_builder::ExcelBuilder

| # | Java method                                              | Rust trait method                              | Status        |
|---|----------------------------------------------------------|------------------------------------------------|---------------|
| 1 | `addContent(Collection<?>, WriteSheet)`                  | `ExcelBuilder::add_content<T,I>(...)`          | [OK-ADAPTED]  |
| 2 | `addContent(Collection<?>, WriteSheet, WriteTable)`      | `ExcelBuilder::add_content_with_table<T,I>(...)` | [OK-ADAPTED] |
| 3 | `fill(Object, FillConfig, WriteSheet)`                   | `ExcelBuilder::fill(...)`                      | [OK-ADAPTED]  |
| 4 | `merge(int, int, int, int)`                              | `ExcelBuilder::merge(...)`                     | [OK-RENAMED]  |
| 5 | `writeContext()`                                         | `ExcelBuilder::write_context()`                | [OK-RENAMED]  |
| 6 | `finish(boolean)`                                        | `ExcelBuilder::finish(on_exception)`           | [OK-RENAMED]  |

### E.1 ExcelBuilderImpl (concrete)

| # | Java method                          | Rust method                                  | Status        |
|---|--------------------------------------|----------------------------------------------|---------------|
| 1 | `ExcelBuilderImpl(WriteWorkbook)`    | `ExcelBuilderImpl::new(writer, path)`        | [OK-ADAPTED]  |
| 2 | `addContent(...)`                    | `ExcelBuilderImpl::add_content(...)`         | [OK-RENAMED]  |
| 3 | `addContent(... WriteTable)`         | `ExcelBuilderImpl::add_content_with_table(...)`| [OK-RENAMED]|
| 4 | `fill(...)`                          | `ExcelBuilderImpl::fill(...)`                 | [OK-RENAMED]  |
| 5 | `merge(...)`                         | `ExcelBuilderImpl::merge(...)`                | [OK-RENAMED]  |
| 6 | `writeContext()`                     | `ExcelBuilderImpl::write_context()`           | [OK-RENAMED]  |
| 7 | `finish(boolean)`                    | `ExcelBuilderImpl::finish(bool)`              | [OK-RENAMED]  |
| 8 | `initStyleMap()` (private)           | `ExcelBuilderImpl::init_style_map()`          | [OK-RENAMED]  |
| 9 | `addStatisticsData()` (private)      | `ExcelBuilderImpl::add_statistics_data()`     | [OK-RENAMED]  |

---

## F. Annotation types (Java `com.alibaba.excel.annotation.*` ↔ Rust `easyexcel_core::annotation::*`)

### F.1 @ExcelProperty ↔ ExcelProperty marker + `#[excel(name, index, order, converter)]`

| # | Java annotation attribute | Rust field                  | Status        |
|---|---------------------------|-----------------------------|---------------|
| 1 | `String[] value()`        | `name: Option<Vec<String>>` | [OK-ADAPTED]  |
| 2 | `int index()`             | `index: Option<i32>`        | [OK-RENAMED]  |
| 3 | `int order()`             | `order: Option<i32>`        | [OK-RENAMED]  |
| 4 | `Class<? extends Converter> converter()` | `converter: Option<Type>` | [OK-ADAPTED] |

### F.2 @ExcelIgnore ↔ `#[excel(ignore)]`

| # | Java | Rust                         | Status |
|---|------|------------------------------|--------|
| 1 | (no fields) | marker field `ignore: bool` | [OK-ADAPTED] |

### F.3 @ExcelIgnoreUnannotated ↔ `#[excel(ignore_unannotated)]`

| # | Java | Rust                                | Status |
|---|------|-------------------------------------|--------|
| 1 | class-level | struct-level `ignore_unannotated: bool` | [OK-ADAPTED] |

### F.4 @DateTimeFormat ↔ `#[excel(format = "...")]`

| # | Java attribute         | Rust field                | Status       |
|---|------------------------|---------------------------|--------------|
| 1 | `String value()`       | `format: Option<String>`  | [OK-RENAMED] |
| 2 | `boolean use1904windowing()` | `use_1904_windowing: Option<bool>` | [OK-RENAMED] |

### F.5 @NumberFormat ↔ `#[excel(format = "...")]`

| # | Java attribute             | Rust field                | Status       |
|---|----------------------------|---------------------------|--------------|
| 1 | `String value()`           | `format: Option<String>`  | [OK-RENAMED] |
| 2 | `RoundingMode roundingMode()` | `rounding_mode: Option<RoundingMode>` | [OK-ADAPTED] |

### F.6 @ColumnWidth ↔ `#[excel(column_width = N)]`

| # | Java    | Rust                              | Status |
|---|---------|-----------------------------------|--------|
| 1 | `int value()` | `column_width: Option<u32>`       | [OK-RENAMED] |

### F.7 @HeadRowHeight ↔ `#[excel(head_row_height = N)]`

| # | Java       | Rust                                 | Status |
|---|------------|--------------------------------------|--------|
| 1 | `short value()` | `head_row_height: Option<u16>`     | [OK-RENAMED] |

### F.8 @ContentRowHeight ↔ `#[excel(content_row_height = N)]`

| # | Java       | Rust                                   | Status |
|---|------------|----------------------------------------|--------|
| 1 | `short value()` | `content_row_height: Option<u16>`    | [OK-RENAMED] |

### F.9 @HeadStyle ↔ `#[excel(head_style(...))]`

| # | Java attribute           | Rust field                  | Status |
|---|--------------------------|-----------------------------|--------|
| 1 | `short dataFormat()`     | `data_format: Option<i16>`  | [OK-RENAMED] |
| 2 | `int font()`             | `font: Option<i32>`         | [OK-RENAMED] |
| 3 | `IndexedColors bgColor()` | `bg_color: Option<IndexedColors>` | [OK-RENAMED] |
| 4 | (none — boolean flags via `hidden`, `locked`, `quotePrefix`, `wrapped`, `shrinkToFit`, `rotation`, `indent`, `horizontalAlignment`, `verticalAlignment`, fill pattern, borders) | derived fields in `HeadStyle` struct | [OK-ADAPTED] |

### F.10 @HeadFontStyle ↔ `#[excel(head_font_style(...))]`

| # | Java attribute                  | Rust field                       | Status |
|---|---------------------------------|----------------------------------|--------|
| 1 | `String fontName()`             | `font_name: Option<String>`      | [OK-RENAMED] |
| 2 | `short fontHeightInPoints()`    | `font_height_in_points: Option<i16>` | [OK-RENAMED] |
| 3 | `boolean bold()`                | `bold: Option<bool>`             | [OK-RENAMED] |
| 4 | `boolean italic()`              | `italic: Option<bool>`           | [OK-RENAMED] |
| 5 | `IndexedColors color()`         | `color: Option<IndexedColors>`   | [OK-RENAMED] |
| 6 | (none — `strikeout`, `charset`, `typeOffset`, `underline`) | derived fields | [GAP] Phase 1 |

### F.11 @ContentStyle ↔ `#[excel(content_style(...))]`

Same shape as @HeadStyle, see F.9. [OK-ADAPTED]

### F.12 @ContentFontStyle ↔ `#[excel(content_font_style(...))]`

Same shape as @HeadFontStyle, see F.10. [OK-ADAPTED]

### F.13 @ContentLoopMerge ↔ `#[excel(content_loop_merge(...))]`

| # | Java attribute        | Rust field                       | Status |
|---|-----------------------|----------------------------------|--------|
| 1 | `int eachRow()`       | `each_row: Option<u32>`          | [OK-RENAMED] |
| 2 | `int columnExtend()`  | `column_extend: Option<u32>`     | [OK-RENAMED] |

### F.14 @OnceAbsoluteMerge ↔ `#[excel(once_absolute_merge(...))]`

| # | Java attribute                | Rust field                          | Status |
|---|-------------------------------|-------------------------------------|--------|
| 1 | `int firstRowIndex()`         | `first_row_index: Option<i32>`      | [OK-RENAMED] |
| 2 | `int lastRowIndex()`          | `last_row_index: Option<i32>`       | [OK-RENAMED] |
| 3 | `int firstColumnIndex()`      | `first_column_index: Option<i32>`   | [OK-RENAMED] |
| 4 | `int lastColumnIndex()`       | `last_column_index: Option<i32>`    | [OK-RENAMED] |

### F.15 GAP markers (Phase 1 to add)

| # | Java equivalent | Rust marker to add             | Derive attr to add          | Phase |
|---|-----------------|--------------------------------|-----------------------------|-------|
| 1 | (none)          | `ExcelImage`                   | `#[excel(image = "...")]`   | 1     |
| 2 | (none)          | `ExcelComment`                 | `#[excel(comment = "...")]` | 1     |
| 3 | (none)          | `ExcelHyperlink`               | `#[excel(hyperlink = "...")]` | 1   |
| 4 | (none)          | `ExcelFormula`                 | `#[excel(formula = "...")]` | 1     |
| 5 | (none)          | `ExcelDataValidation`          | `#[excel(data_validation(...))]` | 1 |
| 6 | (none)          | `ExcelConditional`             | `#[excel(conditional(...))]` | 1   |
| 7 | (none)          | `ExcelFilter`                  | `#[excel(filter)]`          | 1     |

---

## G. com.alibaba.excel.read.listener.ReadListener ↔ easyexcel_core::ReadListener

| # | Java method                                       | Rust trait method                       | Status       |
|---|---------------------------------------------------|-----------------------------------------|--------------|
| 1 | `invoke(T data, AnalysisContext context)`         | `invoke(&mut self, data, ctx)`          | [OK-ADAPTED] |
| 2 | `doAfterAllAnalysed(AnalysisContext context)`     | `do_after_all_analysed(&mut self, ctx)` | [OK-RENAMED] |
| 3 | `hasNext(AnalysisContext context)`                | `has_next(&mut self, ctx)`              | [OK-RENAMED] |
| 4 | `invokeHead(Map, AnalysisContext)`                | `invoke_head(&mut self, head, ctx)`     | [OK-RENAMED] |
| 5 | `onException(Exception, AnalysisContext)`         | `on_exception(&mut self, err, ctx)`     | [OK-RENAMED] |
| 6 | `extra(CellExtra, AnalysisContext)`               | `extra(&mut self, extra, ctx)`          | [OK-RENAMED] |

---

## H. com.alibaba.excel.write.handler.WriteHandler hierarchy ↔ easyexcel_core::WriteHandler

| # | Java interface method                                                  | Rust trait method                         | Status       |
|---|------------------------------------------------------------------------|-------------------------------------------|--------------|
| 1 | `order()` (in `Order` interface)                                       | `order()`                                 | [OK]         |
| 2 | `beforeWorkbookCreate()`                                               | `before_workbook()`                       | [OK-RENAMED] |
| 3 | `afterWorkbookCreate(Workbook)`                                        | `after_workbook()`                        | [OK-RENAMED] |
| 4 | `afterWorkbookDispose(WorkbookWriteContext)`                           | (in `after_workbook()`)                   | [HANDLE]     |
| 5 | `beforeSheetCreate(SheetWriteHandlerContext)`                          | `before_sheet(ctx)`                       | [OK-RENAMED] |
| 6 | `afterSheetCreate(SheetWriteHandlerContext)`                           | `after_sheet(ctx)`                        | [OK-RENAMED] |
| 7 | `beforeSheetDispose(SheetWriteHandlerContext)`                         | (in `after_sheet`?)                       | [HANDLE]     |
| 8 | `afterSheetDispose(SheetWriteHandlerContext)`                          | (in `after_sheet`?)                       | [HANDLE]     |
| 9 | `beforeRowCreate(RowWriteHandlerContext)`                              | `before_row(ctx)`                         | [OK-RENAMED] |
| 10| `afterRowCreate(RowWriteHandlerContext)`                               | `after_row(ctx)`                          | [OK-RENAMED] |
| 11| `beforeRowDispose(RowWriteHandlerContext)`                             | (combined with after_row?)                | [HANDLE]     |
| 12| `afterRowDispose(RowWriteHandlerContext)`                              | (combined with after_row?)                | [HANDLE]     |
| 13| `beforeCellCreate(CellWriteHandlerContext, Cell, Head)`                 | `before_cell(ctx)`                        | [OK-RENAMED] |
| 14| `afterCellCreate(CellWriteHandlerContext, Cell, Head)`                  | `after_cell(ctx)`                         | [OK-RENAMED] |
| 15| `afterCellDataConverted(CellWriteHandlerContext, WriteCellData, Cell, Head)` | (in `after_cell`)                  | [HANDLE]     |
| 16| `afterCellDispose(CellWriteHandlerContext, Cell, Head)`                | (in `after_cell`)                         | [HANDLE]     |

**Phase 2 action items**: keep `WriteHandler` flat for backward compat AND split out sub-traits `WorkbookWriteHandler`, `SheetWriteHandler`, `RowWriteHandler`, `CellWriteHandler`, `MergeHandler`, `ConstraintHandler` mirroring Java sub-interfaces.

---

## I. com.alibaba.excel.write.handler.*StyleStrategy (Java) ↔ easyexcel_writer::style::*

| # | Java class                                       | Rust type                          | Status       |
|---|--------------------------------------------------|------------------------------------|--------------|
| 1 | `AbstractCellStyleStrategy`                      | `AbstractCellStyleStrategy` (trait) | [OK]         |
| 2 | `HorizontalCellStyleStrategy`                    | `HorizontalCellStyleStrategy`       | [OK]         |
| 3 | `AbstractVerticalCellStyleStrategy`              | `AbstractVerticalCellStyleStrategy` | [OK]         |
| 4 | `VerticalCellStyleStrategy`                      | `VerticalCellStyleStrategy`         | [OK]         |
| 5 | `AbstractColumnWidthStyleStrategy`               | `AbstractColumnWidthStyleStrategy`  | [OK]         |
| 6 | `SimpleColumnWidthStyleStrategy`                 | `SimpleColumnWidthStyleStrategy`    | [OK]         |
| 7 | `LongestMatchColumnWidthStyleStrategy`           | `LongestMatchColumnWidthStyleStrategy` | [OK]      |
| 8 | `AbstractRowHeightStyleStrategy`                 | `AbstractRowHeightStyleStrategy`    | [OK]         |
| 9 | `SimpleRowHeightStyleStrategy`                   | `SimpleRowHeightStyleStrategy`      | [OK]         |

---

## J. com.alibaba.excel.exception.* ↔ easyexcel_core::exception::*

| # | Java exception class                       | Rust error variant                          | Status       |
|---|--------------------------------------------|---------------------------------------------|--------------|
| 1 | `ExcelRuntimeException`                    | `ExcelError::Runtime`                       | [OK-ADAPTED] |
| 2 | `ExcelCommonException`                     | `ExcelError::Common`                        | [OK-ADAPTED] |
| 3 | `ExcelAnalysisException`                   | `ExcelError::Analysis`                      | [OK-ADAPTED] |
| 4 | `ExcelAnalysisStopException`               | `ExcelError::AnalysisStop`                  | [OK-ADAPTED] |
| 5 | `ExcelAnalysisStopSheetException`          | `ExcelError::AnalysisStopSheet`             | [OK-ADAPTED] |
| 6 | `ExcelDataConvertException`                | `ExcelError::DataConvert`                   | [OK-ADAPTED] |
| 7 | `ExcelWriteDataConvertException`           | `ExcelError::WriteDataConvert`              | [OK-ADAPTED] |
| 8 | `ExcelGenerateException`                   | `ExcelError::Generate`                      | [OK-ADAPTED] |
| 9 | `ExcelError::Unsupported`                  | `ExcelError::Unsupported(String)`           | [OK]         |
| 10| `ExcelError::Io`                           | `ExcelError::Io(std::io::Error)`            | [OK]         |
| 11| `ExcelError::Format`                       | `ExcelError::Format(String)`                | [OK]         |

---

## K. com.alibaba.excel.cache.* ↔ easyexcel_reader::cache::*

| # | Java class                          | Rust type                          | Status |
|---|-------------------------------------|------------------------------------|--------|
| 1 | `MapCache`                          | `MapCache`                         | [OK]   |
| 2 | `Ehcache`                           | `Ehcache`                          | [OK]   |
| 3 | `XlsCache`                          | `XlsCache`                         | [OK]   |
| 4 | `SimpleReadCacheSelector`           | `SimpleReadCacheSelector`          | [OK]   |
| 5 | `EternalReadCacheSelector`          | `EternalReadCacheSelector`         | [OK]   |
| 6 | `ReadCache` interface               | `ReadCache` trait                  | [OK]   |
| 7 | `ReadCacheSelector` interface       | `ReadCacheSelector` trait          | [OK]   |

---

## L. com.alibaba.excel.converters.Converter ↔ easyexcel_core::converter::Converter

| # | Java method                                                       | Rust method                              | Status       |
|---|-------------------------------------------------------------------|------------------------------------------|--------------|
| 1 | `Class<?> supportJavaTypeKey()`                                   | (encoded as generic param `T`)           | [HANDLE]     |
| 2 | `CellDataTypeEnum supportExcelTypeKey()`                          | `support_excel_type(&self) -> CellDataType` | [OK-RENAMED] |
| 3 | `T convertToJavaData(ReadCellData, ExcelContentProperty, GlobalConfiguration)` | `convert_to_rust_data(&self, ReadConverterContext<T>)` | [OK-ADAPTED] |
| 4 | `T convertToJavaData(ReadConverterContext)`                       | `convert_to_rust_data(&self, ReadConverterContext<T>)` | [OK-ADAPTED] |
| 5 | `WriteCellData convertToExcelData(T, ExcelContentProperty, GlobalConfiguration)` | `convert_to_excel_data(&self, WriteConverterContext<T>)` | [OK-ADAPTED] |
| 6 | `WriteCellData convertToExcelData(WriteConverterContext)`         | `convert_to_excel_data(&self, WriteConverterContext<T>)` | [OK-ADAPTED] |

### L.1 Built-in converter class mapping (60+ types)

| Java class                                        | Rust module                                       | Status |
|---------------------------------------------------|---------------------------------------------------|--------|
| `ByteArrayImageConverter`                         | `converter::bytearray::byte_array_image_converter` | [OK]   |
| `BoxingByteArrayImageConverter`                   | `converter::bytearray::boxing_byte_array_image_converter` | [OK] |
| `BigDecimalBooleanConverter`                      | `converter::bigdecimal::big_decimal_boolean_converter` | [OK] |
| `BigDecimalNumberConverter`                       | `converter::bigdecimal::big_decimal_number_converter` | [OK] |
| `BigDecimalStringConverter`                       | `converter::bigdecimal::big_decimal_string_converter` | [OK] |
| `BigIntegerBooleanConverter`                      | `converter::biginteger::big_integer_boolean_converter` | [OK] |
| `BigIntegerNumberConverter`                       | `converter::biginteger::big_integer_number_converter` | [OK] |
| `BigIntegerStringConverter`                       | `converter::biginteger::big_integer_string_converter` | [OK] |
| `BooleanBooleanConverter`                         | `converter::booleanconverter::boolean_boolean_converter` | [OK] |
| `BooleanNumberConverter`                          | `converter::booleanconverter::boolean_number_converter` | [OK] |
| `BooleanStringConverter`                          | `converter::booleanconverter::boolean_string_converter` | [OK] |
| `ByteBooleanConverter`                            | `converter::byteconverter::byte_boolean_converter` | [OK]   |
| `ByteNumberConverter`                             | `converter::byteconverter::byte_number_converter` | [OK]   |
| `ByteStringConverter`                             | `converter::byteconverter::byte_string_converter` | [OK]   |
| `DateDateConverter`                               | `converter::date::date_date_converter`           | [OK]   |
| `DateNumberConverter`                             | `converter::date::date_number_converter`         | [OK]   |
| `DateStringConverter`                             | `converter::date::date_string_converter`         | [OK]   |
| `DoubleBooleanConverter`                          | `converter::doubleconverter::double_boolean_converter` | [OK] |
| `DoubleNumberConverter`                           | `converter::doubleconverter::double_number_converter` | [OK] |
| `DoubleStringConverter`                           | `converter::doubleconverter::double_string_converter` | [OK] |
| `FileImageConverter`                              | `converter::file::file_image_converter`          | [OK]   |
| `FloatBooleanConverter`                           | `converter::floatconverter::float_boolean_converter` | [OK] |
| `FloatNumberConverter`                            | `converter::floatconverter::float_number_converter` | [OK]  |
| `FloatStringConverter`                            | `converter::floatconverter::float_string_converter` | [OK]  |
| `InputStreamImageConverter`                       | `converter::inputstream::input_stream_image_converter` | [OK] |
| `IntegerBooleanConverter`                         | `converter::integer::integer_boolean_converter` | [OK]   |
| `IntegerNumberConverter`                          | `converter::integer::integer_number_converter`  | [OK]   |
| `IntegerStringConverter`                          | `converter::integer::integer_string_converter`  | [OK]   |
| `LocalDateDateConverter`                          | `converter::localdate::local_date_date_converter` | [OK] |
| `LocalDateNumberConverter`                        | `converter::localdate::local_date_number_converter` | [OK] |
| `LocalDateStringConverter`                        | `converter::localdate::local_date_string_converter` | [OK] |
| `LocalDateTimeDateConverter`                      | `converter::localdatetime::local_date_time_date_converter` | [OK] |
| `LocalDateTimeNumberConverter`                    | `converter::localdatetime::local_date_time_number_converter` | [OK] |
| `LocalDateTimeStringConverter`                    | `converter::localdatetime::local_date_time_string_converter` | [OK] |
| `LongBooleanConverter`                            | `converter::longconverter::long_boolean_converter` | [OK] |
| `LongNumberConverter`                             | `converter::longconverter::long_number_converter` | [OK]  |
| `LongStringConverter`                             | `converter::longconverter::long_string_converter` | [OK]  |
| `ShortBooleanConverter`                           | `converter::shortconverter::short_boolean_converter` | [OK] |
| `ShortNumberConverter`                            | `converter::shortconverter::short_number_converter` | [OK] |
| `ShortStringConverter`                            | `converter::shortconverter::short_string_converter` | [OK] |
| `StringBooleanConverter`                          | `converter::string::string_boolean_converter`    | [OK]   |
| `StringNumberConverter`                           | `converter::string::string_number_converter`     | [OK]   |
| `StringStringConverter`                           | `converter::string::string_string_converter`     | [OK]   |
| `StringErrorConverter`                            | `converter::string::string_error_converter`      | [OK]   |
| `StringImageConverter`                            | `converter::string::string_image_converter`      | [OK]   |
| `UrlImageConverter`                               | `converter::url::url_image_converter`            | [OK]   |

---

## M. Phase tracking summary

| Phase | What gets implemented                                          | New Rust files | Estimated method count |
|-------|----------------------------------------------------------------|---------------|------------------------|
| 1     | ExcelImage/ExcelComment/ExcelHyperlink/ExcelFormula/ExcelDataValidation/ExcelConditional/ExcelFilter markers + derive attrs | +7 marker + +derive attr tests | ~30 |
| 2     | Sub-traits: WorkbookWriteHandler/SheetWriteHandler/RowWriteHandler/CellWriteHandler/MergeHandler/ConstraintHandler + default impls | +5 trait + +5 default impl + +tests | ~50 |
| 3     | Comment read/write, hyperlink, formula, data validation, conditional formatting, auto-filter, split pane, print/header/footer | +5-8 new modules + +tests | ~80 |
| 4     | POI handle + WriteTable overload + ExcelWriter::write_with_table | +1-2 modules + +tests | ~10 |
| 5     | legacy XLS template fill + encryption + image + extra metadata | +2-3 modules + +tests | ~30 |
| 6     | Harden 1:1 test matrix assertions                             | refactor tests | ~50 test changes |
| 7     | Golden JSON alignment                                           | regenerate + +tests | ~20 |
