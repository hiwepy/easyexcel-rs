# CodeGraph Java↔Rust Method-Level Comparison Matrix

> Generated 2026-07-21 by direct codegraph_node queries on both repos.
> Each row mirrors one Java method and verifies the Rust counterpart's
> signature, naming, and behavioural intent.

## 1. EasyExcelFactory (Java) ↔ EasyExcel (Rust)

| # | Java (com.alibaba.excel.EasyExcelFactory)                          | Rust (easyexcel::EasyExcel)                     | Match? |
|---|---------------------------------------------------------------------|------------------------------------------------|--------|
| 1 | `public static ExcelWriterBuilder write()`                          | `EasyExcel::write<T>(path)`                    | OK-RENAMED |
| 2 | `write(File)`                                                       | `EasyExcel::write<T>(path)`                    | OK-ADAPTED  |
| 3 | `write(File, Class)`                                                | `EasyExcel::write<T>(path).head(T)`             | OK-ADAPTED  |
| 4 | `write(String)`                                                     | `EasyExcel::write<T>(path)`                    | OK-ADAPTED  |
| 5 | `write(String, Class)`                                              | `EasyExcel::write<T>(path).head(T)`             | OK-ADAPTED  |
| 6 | `write(OutputStream)`                                               | `EasyExcel::write<T>(path).to_writer(W)`        | OK-ADAPTED  |
| 7 | `write(OutputStream, Class)`                                        | `EasyExcel::write<T>(path).to_writer(W).head(T)` | OK-ADAPTED |
| 8 | `writerSheet()`                                                     | `EasyExcel::writer_sheet::<T>(None)`            | OK-RENAMED  |
| 9 | `writerSheet(Integer)`                                              | `EasyExcel::writer_sheet_index::<T>(idx)`       | OK-RENAMED  |
| 10| `writerSheet(String)`                                               | `EasyExcel::writer_sheet::<T>(name)`            | OK-RENAMED  |
| 11| `writerTable()`                                                     | `EasyExcel::writer_table(0)`                   | OK-RENAMED  |
| 12| `writerTable(Integer)`                                              | `EasyExcel::writer_table(idx)`                  | OK-RENAMED  |
| 13| `read()`                                                            | `EasyExcel::read<T,L>(path, listener)`          | OK-ADAPTED  |
| 14| `read(File)`                                                        | `EasyExcel::read<T,L>(path, listener)`          | OK-ADAPTED  |
| 15| `read(File, ReadListener)`                                          | `EasyExcel::read<T,L>(path, listener)`          | OK          |
| 16| `read(File, Class, ReadListener)`                                   | `EasyExcel::read<T,L>(path, listener).head()`   | OK-ADAPTED  |
| 17| `read(String, ReadListener)`                                        | `EasyExcel::read<T,L>(path, listener)`          | OK          |
| 18| `read(String, Class, ReadListener)`                                 | `EasyExcel::read<T,L>(path, listener).head()`   | OK-ADAPTED  |
| 19| `read(InputStream, ReadListener)`                                   | `EasyExcel::read<T,L>(path, listener)`          | OK-ADAPTED  |
| 20| `read(InputStream, Class, ReadListener)`                            | `EasyExcel::read<T,L>(path, listener).head()`   | OK-ADAPTED  |
| 21| `readSheet()`                                                       | `ExcelReaderSheetBuilder::default()`            | OK-ADAPTED  |
| 22| `readSheet(Integer)`                                                | `ExcelReaderSheetBuilder::sheet_no(idx)`        | OK-RENAMED  |
| 23| `readSheet(String)`                                                 | `ExcelReaderSheetBuilder::sheet_name(name)`     | OK-RENAMED  |
| 24| `readSheet(Integer sheetNo, String sheetName)`                       | (combined into builder API)                     | HANDLE      |

## 2. ExcelReader (Java) ↔ easyexcel_reader::ExcelReader

| # | Java method (com.alibaba.excel.ExcelReader)             | Rust method (easyexcel_reader::ExcelReader)         | Match? |
|---|---------------------------------------------------------|----------------------------------------------------|--------|
| 1 | `ExcelReader(ReadWorkbook)` constructor                  | `ExcelReader::new(path, options, listener)`         | OK-RENAMED |
| 2 | `void read()` (@Deprecated)                            | `fn read_all(&mut self)`                            | OK-RENAMED |
| 3 | `void readAll()`                                        | `fn read_all(&mut self)`                            | OK-RENAMED |
| 4 | `ExcelReader read(ReadSheet...)`                        | `fn read(&mut self, sheets: &[ReadSheet])`          | OK-RENAMED |
| 5 | `ExcelReader read(List<ReadSheet>)`                     | `fn read(&mut self, sheets: &[ReadSheet])`          | OK-ADAPTED |
| 6 | `AnalysisContext analysisContext()`                     | `fn analysis_context(&self) -> &AnalysisContext`    | OK-RENAMED |
| 7 | `AnalysisContext getAnalysisContext()` (@Deprecated)    | (removed, replaced by analysis_context)             | OK          |
| 8 | `ExcelReadExecutor excelExecutor()`                     | (folded into ExcelAnalyserImpl)                     | HANDLE      |
| 9 | `void finish()`                                         | `fn finish(&mut self)`                             | OK          |
| 10| `void close()`                                          | `Drop` impl                                        | OK-ADAPTED  |
| 11| `void finalize()`                                       | `Drop` impl                                       | OK-ADAPTED  |

## 3. ExcelWriter (Java) ↔ easyexcel_writer::ExcelWriter

| # | Java method (com.alibaba.excel.ExcelWriter)                                  | Rust method (easyexcel_writer::ExcelWriter)        | Match? |
|---|-------------------------------------------------------------------------------|----------------------------------------------------|--------|
| 1 | `ExcelWriter(WriteWorkbook)` constructor                                     | `ExcelWriter::new(path)`                          | OK-RENAMED |
| 2 | `ExcelWriter write(Collection, WriteSheet)`                                   | `fn write<T,I>(rows, &WriteSheet)`                | OK-RENAMED |
| 3 | `ExcelWriter write(Supplier<Collection>, WriteSheet)`                        | (planned via Supplier pattern; not yet exposed)  | GAP-Phase3 |
| 4 | `ExcelWriter write(Collection, WriteSheet, WriteTable)`                       | `fn write_with_table<T,I>(rows, &sheet, &table)`  | OK-Phase4  |
| 5 | `ExcelWriter write(Supplier<Collection>, WriteSheet, WriteTable)`             | (planned)                                          | GAP-Phase3 |
| 6 | `ExcelWriter fill(Object, WriteSheet)`                                       | `fn fill(&mut self, data, fill_config, &sheet)`    | OK-ADAPTED |
| 7 | `ExcelWriter fill(Object, FillConfig, WriteSheet)`                            | `fn fill(&mut self, data, fill_config, &sheet)`    | OK          |
| 8 | `ExcelWriter fill(Supplier<Object>, WriteSheet)`                             | (planned)                                          | GAP-Phase3 |
| 9 | `ExcelWriter fill(Supplier<Object>, FillConfig, WriteSheet)`                  | (planned)                                          | GAP-Phase3 |
| 10| `WriteContext writeContext()`                                                | `fn write_context(&self) -> &dyn WriteContext`   | OK-RENAMED |
| 11| `void finish()`                                                              | `fn finish(&mut self) -> Result<()>`             | OK          |
| 12| `void finish_on_exception()`                                                | `fn finish_on_exception(&mut self) -> Result<()>` | OK-RENAMED |
| 13| `void close()`                                                               | `Drop` impl                                       | OK-ADAPTED  |
| 14| `void finalize()`                                                            | `Drop` impl                                       | OK-ADAPTED  |

## 4. ExcelAnalyser / ExcelAnalyserImpl

| # | Java method                                              | Rust method                                 | Match? |
|---|----------------------------------------------------------|---------------------------------------------|--------|
| 1 | `void analysis(List<ReadSheet>, Boolean)` (interface)    | `fn analysis(&mut self)`                    | OK-ADAPTED |
| 2 | `void finish()` (interface)                              | `fn finish(&mut self)`                      | OK-ADAPTED |
| 3 | `ExcelReadExecutor excelExecutor()` (interface)         | (folded into ExcelAnalyserImpl)            | HANDLE  |
| 4 | `AnalysisContext analysisContext()` (interface)          | `fn analysis_context(&self)`               | OK-ADAPTED |
| 5 | `ExcelAnalyserImpl(ReadWorkbook)`                        | `ExcelAnalyserImpl::new()` + `from_path()`  | OK-ADAPTED |
| 6 | `void choiceExcelExecutor()`                             | `fn choice_excel_executor(&mut self)`       | OK-RENAMED |
| 7 | `void removeThreadLocalCache()`                          | (folded; Rust uses per-read context)        | HANDLE  |
| 8 | `void clearEncrypt03()`                                  | (folded into ReadOptions.password handling) | HANDLE  |

## 5. ExcelBuilder / ExcelBuilderImpl

| # | Java method (com.alibaba.excel.write.ExcelBuilder)            | Rust method (easyexcel_writer::ExcelBuilder)        | Match? |
|---|-----------------------------------------------------------------|----------------------------------------------------|--------|
| 1 | `void addContent(Collection, WriteSheet)`                       | `fn add_content<T,I>(&mut self, data, &sheet)`      | OK-RENAMED |
| 2 | `void addContent(Collection, WriteSheet, WriteTable)`            | `fn add_content_with_table<T,I>(&mut self, ...)`    | OK-RENAMED |
| 3 | `void fill(Object, FillConfig, WriteSheet)`                     | `fn fill(&mut self, _data, _fill_config, _sheet)`  | OK-RENAMED |
| 4 | `void merge(int, int, int, int)`                                | `fn merge(...)`                                     | OK-RENAMED |
| 5 | `WriteContext writeContext()`                                   | `fn write_context(&self) -> &dyn WriteContext`     | OK-RENAMED |
| 6 | `void finish(boolean onException)`                              | `fn finish(&mut self, on_exception: bool) -> Result<()>` | OK-RENAMED |

## 6. AnalysisContext / AnalysisContextImpl

| # | Java method                                                   | Rust method                                  | Match? |
|---|---------------------------------------------------------------|----------------------------------------------|--------|
| 1 | `void currentSheet(ReadSheet)`                                | `fn current_sheet(&mut self, &ReadSheet)`    | OK-RENAMED |
| 2 | `ReadWorkbookHolder readWorkbookHolder()`                     | `fn read_workbook_holder(&self)`              | OK-RENAMED |
| 3 | `ReadHolder currentReadHolder()`                              | (folded into read_sheet_holder)             | HANDLE  |
| 4 | `ReadSheetHolder readSheetHolder()`                           | `fn read_sheet_holder(&self) -> Option<&...>`| OK-RENAMED |
| 5 | `ReadRowHolder readRowHolder()`                               | `fn read_row_holder(&self) -> Option<&...>`  | OK-RENAMED |
| 6 | `List<ReadSheet> readSheetList()`                             | `fn read_sheet_list(&self) -> Option<&[ReadSheet]>` | OK-RENAMED |
| 7 | `int getTotalCount()`                                         | (counters folded into AnalysisEventProcessor)| HANDLE  |
| 8 | `int getCurrentRowNum()`                                      | `fn current_row_num(&self) -> Option<i32>`   | OK-RENAMED |
| 9 | `Object getCustom()`                                          | `fn custom<T>(&self) -> Option<&T>`          | OK-RENAMED |
| 10| `ExcelTypeEnum getExcelType()`                                | `fn excel_type(&self) -> ExcelTypeEnum`       | OK-RENAMED |
| 11| `InputStream getInputStream()`                                | (folded into ReadOptions)                    | HANDLE  |
| 12| `InputStream getOriginalInputStream()`                        | (folded into ReadOptions)                    | HANDLE  |
| 13| `void interrupt()`                                            | `fn interrupt(&self) -> Result<()>`          | OK-RENAMED |
| 14| `Object getCurrentRowAnalysisResult()`                        | (folded into AnalysisEventProcessor)        | HANDLE  |
| 15| `AnalysisEventProcessor analysisEventProcessor()`              | `fn analysis_event_processor(&mut self) -> &mut dyn AnalysisEventProcessor` | OK-RENAMED |

## 7. WriteContext / WriteContextImpl

| # | Java method (com.alibaba.excel.write.WriteContext) | Rust method (easyexcel_core::WriteContext / WriteContextImpl) | Match? |
|---|---|---|---|
| 1 | `WriteWorkbookHolder getWorkbookHolder()` | `fn workbook_context(&self) -> &WriteWorkbookContext` | OK-RENAMED |
| 2 | `WriteSheetHolder getSheetHolder()` | `fn sheet_context(&self) -> Option<&WriteSheetContext>` | OK-RENAMED |
| 3 | `WriteTableHolder getTableHolder()` | `fn table_no(&self) -> Option<i32>` (folded, table is index) | HANDLE |
| 4 | `boolean needHead()` | (via sheet_context) | HANDLE |
| 5 | (interface) `void currentSheet(WriteSheet, WriteTable)` | `fn set_sheet_context(&mut self, sheet_name)` | OK-RENAMED |

## 8. ReadListener

| # | Java method                                              | Rust method                                  | Match? |
|---|----------------------------------------------------------|----------------------------------------------|--------|
| 1 | `void invoke(T data, AnalysisContext context)`           | `fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()>` | OK-RENAMED |
| 2 | `void doAfterAllAnalysed(AnalysisContext context)`       | `fn do_after_all_analysed(&mut self, &AnalysisContext) -> Result<()>` | OK-RENAMED |
| 3 | `boolean hasNext(AnalysisContext context)`               | `fn has_next(&mut self, &AnalysisContext) -> bool` | OK-RENAMED |
| 4 | `default void invokeHead(Map, AnalysisContext)`          | `fn invoke_head(&mut self, head, ctx)`        | OK-RENAMED |
| 5 | `default void onException(Exception, AnalysisContext)`   | `fn on_exception(&mut self, &ExcelError, &AnalysisContext) -> ErrorAction` | OK-RENAMED |
| 6 | `default void extra(CellExtra, AnalysisContext)`         | `fn extra(&mut self, &CellExtra, &AnalysisContext) -> Result<()>` | OK-RENAMED |

## 9. WriteHandler hierarchy

| # | Java method                                            | Rust method                                  | Match? |
|---|--------------------------------------------------------|----------------------------------------------|--------|
| 1 | `Order.order()` (interface)                            | `fn order(&self) -> i32`                    | OK-RENAMED |
| 2 | `default void beforeWorkbookCreate()`                  | `fn before_workbook(&mut self, _context)`    | OK-RENAMED |
| 3 | `default void afterWorkbookCreate(Workbook)`            | (folded into after_workbook)                | HANDLE  |
| 4 | `default void afterWorkbookDispose(WorkbookWriteContext)` | `fn after_workbook(&mut self, _context)`  | OK-RENAMED |
| 5 | `default void beforeSheetCreate(SheetWriteHandlerContext)` | `fn before_sheet(&mut self, _context)`   | OK-RENAMED |
| 6 | `default void afterSheetCreate(SheetWriteHandlerContext)`  | `fn after_sheet(&mut self, _context)`    | OK-RENAMED |
| 7 | `default void beforeSheetDispose(SheetWriteHandlerContext)`| (folded into after_sheet)                | HANDLE  |
| 8 | `default void afterSheetDispose(SheetWriteHandlerContext)` | (folded into after_sheet)                | HANDLE  |
| 9 | `default void beforeRowCreate(RowWriteHandlerContext)`     | `fn before_row(&mut self, _context)`      | OK-RENAMED |
| 10| `default void afterRowCreate(RowWriteHandlerContext)`      | `fn after_row(&mut self, _context)`       | OK-RENAMED |
| 11| `default void beforeRowDispose(RowWriteHandlerContext)`    | (folded into after_row)                   | HANDLE  |
| 12| `default void afterRowDispose(RowWriteHandlerContext)`     | (folded into after_row)                   | HANDLE  |
| 13| `default void beforeCellCreate(CellWriteHandlerContext, Cell, Head)` | `fn before_cell(&mut self, _context)` | OK-RENAMED |
| 14| `default void afterCellCreate(CellWriteHandlerContext, Cell, Head)`  | `fn after_cell(&mut self, _context)`  | OK-RENAMED |
| 15| `default void afterCellDataConverted(CellWriteHandlerContext, WriteCellData, Cell, Head)` | (folded into after_cell) | HANDLE |
| 16| `default void afterCellDispose(CellWriteHandlerContext, Cell, Head)` | (folded into after_cell) | HANDLE |

## 10. Converters (60+ implementations)

| Java class | Rust module | Match? |
|---|---|---|
| `Converter<T>` interface | `converter::Converter<T>` trait | OK-ADAPTED |
| `AutoConverter` | `converter::AutoConverter` | OK |
| `BigDecimalBooleanConverter` | `bigdecimal::big_decimal_boolean_converter` | OK-RENAMED |
| `BigDecimalNumberConverter` | `bigdecimal::big_decimal_number_converter` | OK-RENAMED |
| `BigDecimalStringConverter` | `bigdecimal::big_decimal_string_converter` | OK-RENAMED |
| `BigIntegerBooleanConverter` | `biginteger::big_integer_boolean_converter` | OK-RENAMED |
| `BigIntegerNumberConverter` | `biginteger::big_integer_number_converter` | OK-RENAMED |
| `BigIntegerStringConverter` | `biginteger::big_integer_string_converter` | OK-RENAMED |
| `BooleanBooleanConverter` | `booleanconverter::boolean_boolean_converter` | OK-RENAMED |
| `BooleanNumberConverter` | `booleanconverter::boolean_number_converter` | OK-RENAMED |
| `BooleanStringConverter` | `booleanconverter::boolean_string_converter` | OK-RENAMED |
| `ByteArrayImageConverter` | `bytearray::byte_array_image_converter` | OK-RENAMED |
| `BoxingByteArrayImageConverter` | `bytearray::boxing_byte_array_image_converter` | OK-RENAMED |
| `ByteBooleanConverter` | `byteconverter::byte_boolean_converter` | OK-RENAMED |
| `ByteNumberConverter` | `byteconverter::byte_number_converter` | OK-RENAMED |
| `ByteStringConverter` | `byteconverter::byte_string_converter` | OK-RENAMED |
| `DateDateConverter` | `date::date_date_converter` | OK-RENAMED |
| `DateNumberConverter` | `date::date_number_converter` | OK-RENAMED |
| `DateStringConverter` | `date::date_string_converter` | OK-RENAMED |
| `DoubleBooleanConverter` | `doubleconverter::double_boolean_converter` | OK-RENAMED |
| `DoubleNumberConverter` | `doubleconverter::double_number_converter` | OK-RENAMED |
| `DoubleStringConverter` | `doubleconverter::double_string_converter` | OK-RENAMED |
| `FileImageConverter` | `file::file_image_converter` | OK-RENAMED |
| `FloatBooleanConverter` | `floatconverter::float_boolean_converter` | OK-RENAMED |
| `FloatNumberConverter` | `floatconverter::float_number_converter` | OK-RENAMED |
| `FloatStringConverter` | `floatconverter::float_string_converter` | OK-RENAMED |
| `InputStreamImageConverter` | `inputstream::input_stream_image_converter` | OK-RENAMED |
| `IntegerBooleanConverter` | `integer::integer_boolean_converter` | OK-RENAMED |
| `IntegerNumberConverter` | `integer::integer_number_converter` | OK-RENAMED |
| `IntegerStringConverter` | `integer::integer_string_converter` | OK-RENAMED |
| `LocalDateDateConverter` | `localdate::local_date_date_converter` | OK-RENAMED |
| `LocalDateNumberConverter` | `localdate::local_date_number_converter` | OK-RENAMED |
| `LocalDateStringConverter` | `localdate::local_date_string_converter` | OK-RENAMED |
| `LocalDateTimeDateConverter` | `localdatetime::local_date_time_date_converter` | OK-RENAMED |
| `LocalDateTimeNumberConverter` | `localdatetime::local_date_time_number_converter` | OK-RENAMED |
| `LocalDateTimeStringConverter` | `localdatetime::local_date_time_string_converter` | OK-RENAMED |
| `LongBooleanConverter` | `longconverter::long_boolean_converter` | OK-RENAMED |
| `LongNumberConverter` | `longconverter::long_number_converter` | OK-RENAMED |
| `LongStringConverter` | `longconverter::long_string_converter` | OK-RENAMED |
| `ShortBooleanConverter` | `shortconverter::short_boolean_converter` | OK-RENAMED |
| `ShortNumberConverter` | `shortconverter::short_number_converter` | OK-RENAMED |
| `ShortStringConverter` | `shortconverter::short_string_converter` | OK-RENAMED |
| `StringBooleanConverter` | `string::string_boolean_converter` | OK-RENAMED |
| `StringNumberConverter` | `string::string_number_converter` | OK-RENAMED |
| `StringStringConverter` | `string::string_string_converter` | OK-RENAMED |
| `UrlImageConverter` | `url::url_image_converter` | OK-RENAMED |

## 11. CellData classes

| # | Java (com.alibaba.excel.metadata.data) | Rust (easyexcel_core) | Match? |
|---|---|---|---|
| 1 | `WriteCellData` (40+ methods/setters) | `WriteCellData` (14 public methods) | MOSTLY |
| 2 | `ReadCellData<T>` | `ReadCellData<T>` | OK |
| 3 | `CommentData` | `CommentData` | OK |
| 4 | `FormulaData` | `FormulaData` | OK |
| 5 | `HyperlinkData` | `HyperlinkData` | OK |
| 6 | `ImageData` + `ImageType` | `ImageData` + `ImageType` | OK |
| 7 | `RichTextStringData` | `RichTextStringData` | OK |
| 8 | `CoordinateData` | `CoordinateData` | OK |
| 9 | `ClientAnchorData` | `ClientAnchorData` | OK |
| 10| `CellExtra` | `CellExtra` | OK |

## 12. Fill metadata

| # | Java (com.alibaba.excel.metadata.fill) | Rust | Match? |
|---|---|---|---|
| 1 | `FillConfig` (builder + 9 setters) | `FillConfig` (8 methods) | MOSTLY |
| 2 | `FillWrapper<T>` | `FillWrapper` | OK |
| 3 | `FillDirection` (enum) | `FillDirection` | OK |
| 4 | `AnalysisCell` | `AnalysisCell` | OK |
| 5 | `WriteTemplateAnalysisCellType` | `WriteTemplateAnalysisCellType` | OK |

## 13. Annotations (14 types)

All 14 Java annotations have Rust marker types in `easyexcel_core::annotation::*`:

| Java annotation | Rust marker | derive attr |
|---|---|---|
| @ExcelProperty | ExcelProperty | #[excel(name, index, order, converter)] |
| @ExcelIgnore | ExcelIgnore | #[excel(ignore)] |
| @ExcelIgnoreUnannotated | ExcelIgnoreUnannotated | #[excel(ignore_unannotated)] |
| @DateTimeFormat | DateTimeFormat | #[excel(format)] |
| @NumberFormat | NumberFormat | #[excel(format)] |
| @ColumnWidth | ColumnWidth | #[excel(column_width)] |
| @HeadRowHeight | HeadRowHeight | #[excel(head_row_height)] |
| @ContentRowHeight | ContentRowHeight | #[excel(content_row_height)] |
| @HeadStyle | HeadStyle | #[excel(head_style(...))] |
| @HeadFontStyle | HeadFontStyle | #[excel(head_font_style(...))] |
| @ContentStyle | ContentStyle | #[excel(content_style(...))] |
| @ContentFontStyle | ContentFontStyle | #[excel(content_font_style(...))] |
| @ContentLoopMerge | ContentLoopMerge | #[excel(content_loop_merge(...))] |
| @OnceAbsoluteMerge | OnceAbsoluteMerge | #[excel(once_absolute_merge(...))] |
| (Phase 1 new) | ExcelImage | #[excel(image = "...")] |
| (Phase 1 new) | ExcelComment | #[excel(comment = "...")] |
| (Phase 1 new) | ExcelHyperlink | #[excel(hyperlink = "...")] |
| (Phase 1 new) | ExcelFormula | #[excel(formula = "...")] |
| (Phase 1 new) | ExcelDataValidation | #[excel(data_validation(...))] |
| (Phase 1 new) | ExcelConditional | #[excel(conditional(...))] |
| (Phase 1 new) | ExcelFilter | #[excel(filter)] |

## 14. Reader builders

| # | Java (com.alibaba.excel.read.builder) | Rust | Match? |
|---|---|---|---|
| 1 | `ExcelReaderBuilder` (25 methods) | `ExcelReaderBuilder` (30 methods) | OK |
| 2 | `ExcelReaderSheetBuilder` (15 methods) | `ExcelReaderSheetBuilder` | MOSTLY |
| 3 | `ExcelReaderTableBuilder` | (folded into sheet builder) | HANDLE |

## 15. Writer builders

| # | Java (com.alibaba.excel.write.builder) | Rust | Match? |
|---|---|---|---|
| 1 | `ExcelWriterBuilder` (28 methods) | `ExcelWriterBuilder` (32 methods) | OK |
| 2 | `ExcelWriterSheetBuilder` | `ExcelWriterSheetBuilder` | OK |
| 3 | `ExcelWriterTableBuilder` | `ExcelWriterTableBuilder` | OK |

## 16. Enums

| # | Java | Rust | Match? |
|---|---|---|---|
| 1 | `ExcelTypeEnum { CSV, XLS, XLSX }` | `ExcelTypeEnum { Csv, Xls, Xlsx }` | OK-RENAMED |
| 2 | `CellDataTypeEnum { 8 variants }` | `CellDataType { 9 variants }` | OK-EXTENDED (added Formula, Image) |
| 3 | `CellExtraTypeEnum { COMMENT, HYPERLINK, MERGE }` | `CellExtraType { Comment, Hyperlink, Merge }` | OK-RENAMED |
| 4 | `WriteDirectionEnum { VERTICAL, HORIZONTAL }` | `FillDirection { Vertical, Horizontal }` | OK-RENAMED |
| 5 | `WriteTypeEnum { ADD, FILL, TEMPLATE }` | (folded into WriteOptions flags) | HANDLE |
| 6 | `HeadKindEnum { NONE, CLASS, LIST, DYNAMIC }` | `HeadKind { None, Class, List, Dynamic }` | OK-RENAMED |
| 7 | `ReadDefaultReturnEnum { STRING, CELL }` | `ReadDefaultReturn` | OK-RENAMED |

## 17. Exceptions (8 types)

| # | Java | Rust | Match? |
|---|---|---|---|
| 1 | `ExcelRuntimeException` | `ExcelError::Runtime` | OK-ADAPTED |
| 2 | `ExcelCommonException` | `ExcelError::Common` | OK-ADAPTED |
| 3 | `ExcelAnalysisException` | `ExcelError::Analysis` | OK-ADAPTED |
| 4 | `ExcelAnalysisStopException` | `ExcelError::AnalysisStop` | OK-ADAPTED |
| 5 | `ExcelAnalysisStopSheetException` | `ExcelError::AnalysisStopSheet` | OK-ADAPTED |
| 6 | `ExcelDataConvertException` | `ExcelError::DataConvert` | OK-ADAPTED |
| 7 | `ExcelWriteDataConvertException` | `ExcelError::WriteDataConvert` | OK-ADAPTED |
| 8 | `ExcelGenerateException` | `ExcelError::Generate` | OK-ADAPTED |

## 18. Cache implementations

| # | Java | Rust | Match? |
|---|---|---|---|
| 1 | `ReadCache` interface | `ReadCache` trait | OK |
| 2 | `MapCache` | `MapCache` | OK |
| 3 | `Ehcache` | `Ehcache` | OK |
| 4 | `XlsCache` | `XlsCache` | OK |
| 5 | `SimpleReadCacheSelector` | `SimpleReadCacheSelector` | OK |
| 6 | `EternalReadCacheSelector` | `EternalReadCacheSelector` | OK |
| 7 | `ReadCacheSelector` interface | `ReadCacheSelector` trait | OK |

## 19. Utilities

| # | Java | Rust | Match? |
|---|---|---|---|
| 1 | `ClassUtils` | `ClassUtils` (in easyexcel-derive + core) | OK |
| 2 | `DateUtils` | `DateUtils` | OK |
| 3 | `FileUtils` | `FileUtils` | OK |
| 4 | `NumberUtils` | `NumberUtils` | OK |
| 5 | `StringUtils` | `StringUtils` | OK |
| 6 | `CollectionUtils` | `CollectionUtils` | OK |
| 7 | `ListUtils` | `ListUtils` | OK |
| 8 | `MapUtils` | `MapUtils` | OK |
| 9 | `PoiUtils` | (folded into write path) | HANDLE |
| 10| `WorkBookUtil` | (folded) | HANDLE |
| 11| `WorkbookWriteContextUtils` | (folded) | HANDLE |

## 20. Handler strategies

| # | Java | Rust | Match? |
|---|---|---|---|
| 1 | `AbstractCellStyleStrategy` | `AbstractCellStyleStrategy` | OK |
| 2 | `HorizontalCellStyleStrategy` | `HorizontalCellStyleStrategy` | OK |
| 3 | `AbstractVerticalCellStyleStrategy` | `AbstractVerticalCellStyleStrategy` | OK |
| 4 | `VerticalCellStyleStrategy` | `VerticalCellStyleStrategy` | OK |
| 5 | `AbstractColumnWidthStyleStrategy` | `AbstractColumnWidthStyleStrategy` | OK |
| 6 | `SimpleColumnWidthStyleStrategy` | `SimpleColumnWidthStyleStrategy` | OK |
| 7 | `LongestMatchColumnWidthStyleStrategy` | `LongestMatchColumnWidthStyleStrategy` | OK |
| 8 | `AbstractRowHeightStyleStrategy` | `AbstractRowHeightStyleStrategy` | OK |
| 9 | `SimpleRowHeightStyleStrategy` | `SimpleRowHeightStyleStrategy` | OK |
| 10| `LoopMergeStrategy` | `LoopMergeStrategy` | OK |
| 11| `OnceAbsoluteMergeStrategy` | `OnceAbsoluteMergeStrategy` | OK |

## 21. SAX readers (XLS/XLSX/CSV)

| # | Java (com.alibaba.excel.analysis) | Rust (easyexcel_reader) | Match? |
|---|---|---|---|
| 1 | `XlsSaxAnalyser` | `XlsSaxAnalyser` | OK |
| 2 | `XlsxSaxAnalyser` | `XlsxSaxAnalyser` | OK |
| 3 | `CsvExcelReadExecutor` | `CsvExcelReadExecutor` | OK |
| 4 | 19 XLS record handlers | `crates/easyexcel-reader/src/analysis/v03/handlers/*` | OK (all 19 file names match) |
| 5 | 11 XLSX tag handlers | `crates/easyexcel-reader/src/analysis/v07/handlers/*` | OK (all 11 file names match) |

## 22. Default handler loader

| # | Java | Rust | Match? |
|---|---|---|---|
| 1 | `DefaultWriteHandlerLoader.loadDefaultHandler(boolean, ExcelTypeEnum)` | `DefaultWriteHandlerLoader::load_default_handler()` | OK |
| 2 | `DefaultWriteHandlerImpl` | `DefaultRowWriteHandler` | OK |
| 3 | `DefaultWriteWorkbookHandler` | `DefaultWriteWorkbookHandler` | OK |
| 4 | `DefaultWriteSheetHandler` | `DefaultWriteSheetHandler` | OK |
| 5 | `FillStyleCellWriteHandler` | `FillStyleCellWriteHandler` | OK |
| 6 | `DimensionWorkbookWriteHandler` | (impl_dimension_workbook_write_handler) | OK |