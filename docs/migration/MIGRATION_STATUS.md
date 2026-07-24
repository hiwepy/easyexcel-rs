# easyexcel-rust Migration Status Tracker

> 迁移进度追踪。以 `docs/migration/file-map.csv` + `xtask migration-audit` 为准。
> **不以「测试绿」替代账本 `complete`。**

## 0. Baseline（Phase 0，2026-07-23）

| Metric | Value |
|--------|-------|
| Java 基线 | EasyExcel **4.0.3** @ `3afdea9d` |
| Rust HEAD | 工作树（本轮结构对齐） |
| Java main 文件（excl package-info） | **325** |
| file-map 行数 | **325** |
| rust 目标文件缺失 | **0**（含 1:1 路径 shim / re-export） |
| Workspace | `crates/*` + `easyexcel-web/{axum,actix}` + `easyexcel-demo/*` + `xtask` |
| JSON | Jackson/Fastjson → `serde` / `serde_json`（`ExcelDownloadErrorBody`） |
| Web | Spring Boot → `easyexcel-web-axum`；Quarkus → `easyexcel-web-actix` |

## 1. 本轮已完成

- [x] 生成 `docs/migration/file-map.csv`
- [x] 落地 `xtask`（`migration-audit` / `migration-audit-strict`）
- [x] 补齐 37 个「路径级」缺失文件（不删既有 `enum_*.rs` / Builder 实现，仅 shim + 新类型）
- [x] 新增 `Font` / `CellData` / `DataFormatData` / `ReadBasicParameter` / `Empty`
- [x] `easyexcel-web-axum` + `easyexcel-web-actix` + 五个 demo crate
- [x] 修复 `enums.rs`/`enums/mod.rs`、`support.rs`/`support/mod.rs` 双模块冲突（现代 `foo.rs + foo/`）
- [x] `cargo check --workspace` 通过

### 1.1 Reader 语义迁移（2026-07-24）

- [x] `ExcelReader.read(ReadSheet...)` 拒绝空列表，不再错误地读取默认 Sheet
- [x] 多 Sheet 参数按工作簿顺序匹配，不再只读取参数列表最后一项
- [x] 名称选择保留 Java `sheetNo == null` 与显式 `sheetNo == 0` 的区别
- [x] `ReadSheet.copyBasicParameter` 只复制读取参数，不再错误覆盖 Sheet 身份
- [x] Sheet 级 `headRowNumber` / `useScientificFormat` 进入实际读取配置
- [x] 主门面及兼容 `ExcelReaderBuilder` 均可保存并顺序调用多个 Listener
- [x] `ExcelReaderSheetBuilder` 通过独占借用绑定 Reader，并提供 Rust 同形
  `do_read` / `do_read_sync`；同步读取会在既有 Listener 之后追加收集器
- [x] 兼容 `ExcelReaderBuilder.do_read_all_sync` 返回真实转换行；注册过的
  Listener 仍先于同步收集器执行
- [x] `charset`、`ignoreEmptyRow`、`customObject`、`password`、
  `extraRead`、`readDefaultReturn` 均写入实际 `ReadOptions`，不再只是 API 外壳

本轮只为上述行为补充了 `test_evidence`，相关账本项仍保持
`in_progress`；在完整类职责和 Java 对照测试闭环前不升级为
`complete`。

### 1.2 Writer Builder 语义迁移（2026-07-24）

- [x] `write/builder/excel_writer_builder.rs` 从文档空模块替换为真实
  `WriteWorkbook → ExcelWriter` Builder
- [x] `write/builder/excel_writer_sheet_builder.rs` 从文档空模块替换为真实
  Writer 所有权、Sheet 参数和 `do_write` 生命周期
- [x] Workbook 与 Sheet 级 `register_write_handler` 均进入实际 Handler
  执行链，并在第一次写入前统一按 `order()` 排序
- [x] `WriteWorkbook.set_file`、模板文件和模板字节 setter 保存真实值，
  不再返回固定 `None` 或静默 no-op
- [x] `ExcelTypeEnum` 进入 `WriteOptions` 和 stateful/one-shot/stream
  后端选择，能够真正覆盖输出文件扩展名
- [x] Java `file(OutputStream)` 以 Rust 类型化
  `ExcelWriterOutputStreamBuilder<W>` 落地，真实写入调用方持有的流，
  并保留 `autoCloseStream` 生命周期语义
- [x] `sheet().table().doWrite(...)` 交接父 Writer/Sheet 所有权，
  进入真实 `write_with_table → finish` 链路；Supplier 形式由
  `do_write_with` 延迟求值
- [x] 公开门面 `sheet(...).do_fill(...)` 同时接受 `TemplateData` 与
  `FillWrapper`，新增显式 `FillConfig` 和 Supplier 形式，并真实贯通
  `direction / forceNewRow / autoStyle`
- [x] 主门面新增 `excel_type`、`use_default_style`、
  `automatic_merge_head` Rust 链式方法

上述 Builder 映射文件已补真实执行测试，但兼容模块的完整方法职责
仍需继续对照所有 inherited 参数及跨 crate 的 `doFill` 构造桥，因此账本
暂时保持 `in_progress`。

### 1.3 FillConfig 完整参数语义（2026-07-24）

- [x] Java `direction` 映射为 `WriteDirection::{Vertical, Horizontal}`
- [x] `forceNewRow` 保持默认 `false` 并进入模板行扩展算法
- [x] `autoStyle` 默认 `true`，并真实传入模板样式继承逻辑
- [x] `init()` 具备幂等生命周期状态，不再只有单字段简化配置
- [x] `easyexcel-writer` 参数测试与 `easyexcel-template` 执行器传播测试通过

`FillConfig.java` 的独立文件职责已经闭环，可在账本中作为首个
`complete` 项；这只表示该文件职责完成，不代表 Writer Builder 或
整个 Java → Rust 迁移完成。

### 1.4 WriteBasicParameter nullable 继承（2026-07-24）

- [x] `relativeHeadRowIndex`、`needHead`、`useDefaultStyle`、
  `automaticMergeHead`、`orderByIncludeColumn` 改为 nullable 显式覆盖
- [x] include/exclude 的索引与字段名集合保留 `None`（继承）和
  `Some(empty)`（显式清空）的差异
- [x] Workbook 的有效参数在创建 Sheet 时继承，不再被 Sheet 默认值覆盖
- [x] Table 未配置时继承 Sheet；显式配置为 Java 默认值时仍能反向覆盖父级
- [x] 自定义 Converter 按 Workbook → Sheet → Table 顺序合并，而不是替换父级
- [x] 实际 XLSX 测试覆盖表头开关、默认粗体样式和列选择的继承/覆盖
- [x] `AbstractWriteHolder` 不再固定 `ignore=false`，真实执行字段名/列号
  include/exclude 判定，并保留空集合清除父级配置的语义

对应映射已补行为证据，但 `AbstractWriteHolder` 的注解 Handler 生成与
own/parent 分层执行链仍需单独对照，因此相关账本项保持 `in_progress`。

### 1.5 WriteHandler 排序与 NotRepeatExecutor 去重（2026-07-24）

- [x] `WriteHandler` trait object 可显式暴露可选 `NotRepeatExecutor` 能力
- [x] 所有 stateful、XLSX、XLS、CSV 写入口先按 `order()` 稳定排序
- [x] 相同 `uniqueValue()` 只保留排序后的第一个处理器
- [x] 未实现 `NotRepeatExecutor` 的同类型处理器仍会全部执行
- [x] 实际 XLSX 生命周期测试证明低 `order` 实例胜出且重复实例不执行

这补齐了 Java `AbstractWriteHolder.sortAndClearUpHandler` 的排序/去重核心
语义。父 Holder 链、默认 Handler Loader 及运行期共享 Handler 实例已在
1.6 闭环；注解生成 Handler 的可观察文件语义已在 1.7 继续闭环，但完整
Handler 上下文等剩余职责仍需继续逐项核对，因此
`AbstractWriteHolder` 与 Handler 系列仍保持 `in_progress`。

### 1.6 WriteHandler 生命周期与默认链路（2026-07-24）

- [x] `WriteHandler` 执行面补齐 Java 的 Workbook、Sheet、Row、Cell
  `beforeCreate / afterCreate / afterDispose` 生命周期，并保留旧 Rust
  回调作为向后兼容转发
- [x] Cell 在写出前按 Java 顺序执行
  `beforeCellCreate → afterCellCreate → afterCellDataConverted`，写出后执行
  `afterCellDispose`
- [x] Workbook、Sheet、Row、Cell 专用 Trait 不再维护一套不会被写引擎调用的
  假回调；它们变为 `WriteHandler` 的类型标记，真实逻辑只有一个执行入口
- [x] Sheet/Table 自有 Handler 在父 Holder Handler 之前合并；相同
  `order + uniqueValue` 时由更具体的 Holder 胜出，符合 Java 稳定排序后去重
- [x] `DefaultWriteHandlerLoader` 按 XLSX/XLS/CSV 和 `useDefaultStyle`
  返回 Java 对应 Handler 集合
- [x] 状态型 `ExcelWriterBuilder` 与便捷 `EasyExcel::write` 均把默认 Handler
  装入真实执行链，不再只是可独立调用但不生效的类
- [x] 状态型 Writer 按 Workbook、Sheet、Table 保存独立 Handler 链；有效链为
  `Table own + Sheet own + Workbook parent`，克隆的是共享句柄，回调始终落到
  同一个有状态 Handler 实例
- [x] 新 Sheet 只补跑 Sheet own 的 Workbook 回调；新 Table 只补跑 Table own
  的 Workbook/Sheet 回调；行列回调使用当前 Holder 的完整有效链
- [x] 同一 Sheet 的多个 Table 分别保存 schema 和 options，不会复用前一个
  Table 的结构；每个新 Table 仅初始化一次自己的表头，重复写同一 Table
  不会重复表头
- [x] XLSX ZIP 保留模板与 XLS BIFF 保留模板的一次性/状态型追加路径均执行
  Row/Cell 完整生命周期；Handler 修改后的值进入真实文件，`skip` 会省略物理
  Cell，不再出现普通写入生效而模板写入静默失效

以上由 `stateful_sheet_handlers_are_isolated_and_reused_by_holder`、
`table_holder_runs_supplementary_callbacks_then_own_parent_row_chain` 和
`multiple_tables_keep_independent_schema_options_and_single_head`，以及
`template_zip_path_runs_row_cell_lifecycle_and_applies_mutation_and_skip` 验证。
Handler 上下注解收集、完整 POI 对象上下文等职责仍未全部对齐，因此本节涉及的
Handler、Holder 和 Loader 账本项继续保持 `in_progress`，不能据此宣称 Writer
或整体迁移完成。

### 1.7 注解生成 Handler 的文件语义（2026-07-24）

Java `AbstractWriteHolder.initAnnotationConfig` 会根据注解创建样式、列宽、
行高、循环合并和绝对合并 Handler。Rust 不伪造只存在于类型层、却不进入
写引擎的 Handler 对象，而是由 `#[derive(ExcelRow)]` 生成静态
`ExcelWriteMetadata`，在 XLSX、XLS 和模板保留写入路径中执行相同的文件语义：

- [x] 类型级和字段级样式、字体、列宽、表头行高、内容行高进入真实写入路径
- [x] `ContentLoopMerge` 和 `OnceAbsoluteMerge` 均写出真实合并区域
- [x] `OnceAbsoluteMerge` 坐标始终是 Sheet 绝对坐标，不再随模板追加行偏移
- [x] 自定义行高 Handler 在注解之后生效，可覆盖注解生成的行高
- [x] 同一 Sheet 新建第二个 Table 时，会初始化该 Table 自有的注解列宽和
  绝对合并，不会错误复用前一个 Table 的元数据
- [x] ZIP 模板保留路径会在不丢失原包部件的情况下写入注解列宽、行高和绝对合并
- [x] ZIP 模板新增 Cell 会把类型级/字段级字体、填充、边框、对齐、数字格式以及
  Style Handler 结果编译并合并到既有 `styles.xml`；字体、填充、边框、自定义
  `numFmt` 和 `cellXfs` 索引全部重映射，模板原有 Cell 的 `s` 索引保持有效
- [x] 等价字体、填充、边框、数字格式和 XF 在后续批次复用，避免分页写入不断
  膨胀样式表，对齐 Java `WriteWorkbookHolder.cellStyleIndexMap` 的缓存语义
- [x] `OnceAbsoluteMergeStrategy` 对负索引返回格式错误，不生成无效区域

对应实文件测试包括
`annotation_dimensions_apply_field_type_and_explicit_precedence`、
`custom_row_height_handler_overrides_annotation_height`、
`multiple_tables_keep_independent_schema_options_and_single_head` 和
`template_annotation_layout_stays_absolute_and_preserves_package`。

完整 POI Handler 上下文及 Java 全部模板夹具仍未对齐。因此上述账本项仍保持
`in_progress`，不能由本节推导整个 Writer 已迁移完成。

### 1.8 Handler Context 阶段语义与真实工具链（2026-07-24）

- [x] `WriteRowContext` 补齐 Java `relativeRowIndex`，XLSX、XLS、CSV 和模板
  路径分别传递表头层级或内容数据序号
- [x] `WriteCellContext` 补充静态 Column/Head、转换前值、`cellDataList`、
  `firstCellData`、目标 Cell 类型和 `ignoreFillStyle`
- [x] `beforeCellCreate` / `afterCellCreate` 阶段的转换数据保持为空；内容 Cell
  转换完成后再填充列表和目标类型，随后触发 `afterCellDataConverted`
- [x] `#[derive(ExcelRow)]` 为每个字段生成 Converter 前默认编码值、Converter
  输出值及声明字段类型；`originalValue/originalFieldClass` 在
  `beforeCellCreate` / `afterCellCreate` 保持空，到转换阶段才激活
- [x] XLSX、XLS、CSV、ZIP 模板和 BIFF 模板路径均传递 Converter 前后两套值，
  不再用转换后的 `CellValue` 复制一份冒充 `originalValue`
- [x] 表头路径与 Java 一致，不再错误触发 `afterCellDataConverted`；但在
  `afterCellDispose` 可读取已生成的表头 `firstCellData`
- [x] `ignoreFillStyle` 在普通 XLSX、XLS 及 ZIP 模板路径真实抑制注解和
  Style Handler 样式，不只是 Context 上一个未使用的布尔字段
- [x] BIFF8 表头补齐 Row 生命周期，动态多级表头按 level 传递相对行索引
- [x] Java 包路径下四个 `*WriteHandlerContext` 不再是脱离执行链的包装对象，
  已改为真实 runtime Context 的类型别名
- [x] `WriteHandlerUtils` 删除 `Option<()> + AtomicU32` 计数器假实现，改为创建
  真实 Context、分发生命周期并传播 Handler 错误；writer 执行入口直接调用它
- [x] 新增后端中立的 `WriteWorkbookHolderView`、`WriteSheetHolderView`、
  `WriteTableHolderView`；Workbook/Sheet/Row/Cell 回调可读取输出路径、已解析
  Sheet 名称/序号、当前物理行、`hasData` 和可选 Table 序号
- [x] Holder 视图已接入 stateful 与一次性 XLSX、XLS、CSV、ZIP/BIFF 模板
  执行链；Table 写入时 Workbook/Sheet 父 Handler 与 Table 自有 Handler
  共享同一个真实 active-table 视图，非 Table 写入保持 `None`
- [x] 未知状态不再用空路径、`0` 或伪 POI 对象填充；仅旧兼容构造器产生的
  非 writer Context 会明确返回 `None`
- [x] `WriteRowContext.row()` / `WriteCellContext.cell()` 不再返回 Context
  自身冒充 POI 对象，改为明确的 `WriteRowHandle` / `WriteCellHandle`
- [x] 逻辑 Handle 支持回调链内可见的 Cell 值覆盖、跳过、样式覆盖及 Row
  高度覆盖；writer 在逻辑 `afterCellDispose` / `afterRowDispose` 后由实际
  XLSX、XLS、CSV 或 ZIP 模板后端提交修改

`handler_context_matches_java_conversion_stages_and_ignore_fill_style` 使用实际
XLSX 验证阶段字段及无 `s` 样式索引；
`handler_context_exposes_real_pre_converter_value_across_write_backends` 覆盖
XLSX、XLS、CSV 和 ZIP 模板路径，并校验 Workbook 文件、Sheet 名称/序号和
当前行；`table_holder_runs_supplementary_callbacks_then_own_parent_row_chain`
进一步校验 Table 写入的 Workbook/Sheet/Row/Cell Holder 视图。facade 的字段
Converter 测试同时验证 derive 生成的 `"source"` 原值、`String` 字段类型及
`"field:source"` 转换结果。
`logical_row_and_cell_handles_commit_real_backend_mutations` 从生成文件重新读取
XLSX、XLS、CSV 和 ZIP 模板，验证 `afterCellDispose` 改值会真实落盘，并验证
XLSX/模板的 Cell 样式与 Row 高度覆盖；第二个 Handler 还能看到前一个 Handler
修改后的逻辑 Cell 值。
`utilities_dispatch_real_handler_contexts` 验证工具方法分发真实 Handler。

仍未完成：Java Holder 暴露的完整 POI `Workbook/Sheet/Row/Cell` API 没有直接
等价物；当前契约覆盖 EasyExcel 生产 Handler 实际依赖的行列位置、改值、跳过、
样式和行高，但尚未覆盖任意 POI 扩展操作，BIFF 模板的回调式样式/行高覆盖也未
完成。Rust 的 `originalValue` 是转换前字段的后端中立 `CellValue` 表示，并以
`original_field_type` 保留声明类型，而非暴露可变的任意 Rust 对象引用。因此
相关 Context/Holder 账本项继续保持 `in_progress`，不能宣称 POI 对象级迁移完成。

### 1.9 WriteContext 终结生命周期（2026-07-24）

- [x] 删除 `finish_write_context` 无条件返回 `Ok(())` 的静默空实现
- [x] 新增 `WriteContextLifecycle`，只有实际拥有 writer 资源的适配器才能实现
  `finish_context(on_exception)`；纯路径/Holder 元数据上下文不会冒充可保存对象
- [x] `ExcelBuilderImpl` 的 `finish(boolean)` 通过 `finish_write_context`
  动态委派到真实 `ExcelWriter::finish` / `finish_on_exception`
- [x] 实际链路保留幂等终结、`afterWorkbookDispose`、XLSX/XLS/CSV/模板保存、
  `writeExcelOnException` 和 `autoCloseStream` 语义，错误继续向调用方传播
- [x] `finish_write_context_dispatches_to_resource_owner` 验证 lifecycle 动态分发；
  `excel_builder_impl_delegates_add_content_and_finish` 验证重复终结只写出一次真实文件

Rust 因所有权边界将 Java 单一 `WriteContext` 拆为只读上下文和资源生命周期能力，
但不再提供“调用成功却没有保存/关闭任何资源”的假 API。完整 POI 加密清理和
ThreadLocal 缓存清理由 Rust 后端的资源模型替代；相关账本仍保持 `in_progress`，
直到所有终结异常组合都有 Java 对照证据。

### 1.10 FileUtils 可变配置与临时目录语义（2026-07-24）

- [x] `tempFilePrefix`、`poiFilesPath`、`cachePath` 从只允许初始化一次且静默忽略
  后续 setter 的 `OnceLock<T>`，改为可并发读写的进程级配置
- [x] setter 在 getter 已经读取过默认值后仍可真实覆盖，符合 Java 可变静态字段
- [x] `createCacheTmpFile` 在当前 `cachePath` 下创建真实临时文件，不再忽略配置
- [x] `createPoiFilesDirectory` 创建并保留当前 `poiFilesPath`，不再返回随即销毁、
  退出函数便删除目录的短生命周期 `TempDir`
- [x] `configured_paths_can_be_replaced_after_first_read` 覆盖重复配置、目录创建及
  缓存文件父目录

`FileUtils` 的其他职责仍需继续核对输入流关闭策略和异常类型，因此账本保持
`in_progress`。`WorkBookUtil` 的 `None`/空操作占位已在 1.12 节移除，但
CSV 元数据对象和 XLS 模板新建工作表仍需继续补齐。

### 1.11 BeanMapUtils 编译期对象映射（2026-07-24）

- [x] 删除 `BeanMapUtils.create(&dyn Any) -> None` 的类型擦除占位
- [x] 新增实际 `BeanMap`，通过 `ExcelRow::schema` 与生产写链使用的
  `to_row_with_converters` 构建字段名 → `CellValue` 映射
- [x] 保留 derive 生成的声明字段类型，提供 Java
  `BeanMap.getPropertyType(fieldName)` 的后端中立等价查询
- [x] 自定义 `ExcelRow` 返回的 schema/value 数量不一致时返回明确格式错误，
  不静默截断字段
- [x] Java `Write#simpleWrite1` 的 Rust 对照测试已从手写 `HashMap` stand-in
  改为调用真实 `bean_map_utils::create`

Rust 不支持 CGLIB 对原对象字段的运行期反射写回；读模型构建由 derive 生成的
`ExcelRow::from_row` 完成。因此本项证明了写入/填充所需的字段读取语义，不代表
Java BeanMap 的任意运行期变异能力完整迁移，账本继续保持 `in_progress`。

### 1.12 WorkBookUtil 后端创建契约（2026-07-24）

- [x] 删除 `create_work_book/create_sheet/create_row/create_cell` 固定返回
  `None` 和 `fill_data_format` 无条件成功的占位实现
- [x] 新增 `WorkBookCreator`、`SheetCreator`、`RowCreator`、`CellCreator`
  后端契约，工具函数真实委派并传播后端错误
- [x] XLSX 主写入链通过该契约创建命名工作表，并在写值前创建、校验逻辑
  Row/Cell；保留 `rust_xlsxwriter` 的常量内存工作表模式
- [x] BIFF8 一次性及有状态写入通过相同契约创建工作表，表头和数据单元格通过
  实际 `Biff8Sheet::set` 落盘，并校验 65,536 × 256 格式上限
- [x] `fill_data_format` 在 `WriteCellData` 中创建缺失样式和
  `DataFormatData`，按 Java 逻辑选择显式/默认格式且不覆盖已有格式
- [x] `workbook_util_creator_chain_materializes_real_xlsx_cells` 生成真实 XLSX 后
  用 Calamine 回读 `(2,4)` 单元格；核心测试覆盖委派链和格式不覆盖语义

本节已证明这些方法不再是假 API，但 `WorkBookUtil` 账本仍保持
`in_progress`：CSV 对象模型已在 1.13 节接入生产写链，但 XLS 模板路径仍不能
创建模板中不存在的工作表。`WriteCellData` 的运行期格式贯通已在 1.14 节完成。

### 1.13 CSV Workbook/Sheet/Row/Cell 对象模型（2026-07-24）

- [x] 删除 `CsvWorkbook`、`CsvSheet`、`CsvRow`、`CsvCell`、
  `CsvCellStyle`、`CsvDataFormat`、`CsvRichTextString` 七个 `()` 类型别名
- [x] `CsvWorkbook` 保存 locale、日期窗口、科学计数法、charset、BOM、
  单工作表、样式表和数据格式表；重复创建 Sheet 返回明确错误
- [x] `CsvSheet` 强制按顺序创建 Row，并提供有界行缓存、已刷盘行错误和
  stateful append 起始行语义
- [x] `CsvRow/CsvCell` 保存真实行列、稀疏单元格、类型化 `CellValue`、
  日期/数字补充类型、公式、富文本和样式，最终构建稠密 CSV record
- [x] `CsvDataFormat` 复用内建格式索引，自定义格式从 82 开始注册并稳定回查
- [x] 生产 `append_csv_records` 通过 `WorkBookUtil` 创建真实 CSV
  Workbook/Sheet/Row/Cell，再交给 `csv::Writer` 流式输出
- [x] 核心对象测试覆盖类型值、稀疏列、单 Sheet、行顺序和重复 Cell；
  writer 的 10 个 CSV 测试继续覆盖 BOM、charset、动态表头、有状态批次、
  Handler 生命周期和 I/O 错误

Java CSV 类为了实现完整 POI 接口包含大量有意空操作；Rust 没有复制这些无效果
方法，而是实现生产链真正使用的状态和约束。因此七项账本仍为 `in_progress`，
直到与 Java CSV 日期/数字格式化、行缓存刷盘边界的 golden tests 完整对齐。

### 1.14 Converter → WriteCellData 运行期样式链（2026-07-24）

- [x] `Converter<T>::convert_to_excel_data` 从错误的 `CellValue` 返回值改为
  Java 对齐的 `WriteCellData`，全局 `ConverterRegistry` 的类型擦除链同步保留
  公式、超链接、批注、图片、样式和运行期数据格式
- [x] `ExcelRow::to_excel_write_row` 的生产写入结果改为
  `Vec<WriteCellData>`；`to_row/to_row_with_converters` 继续返回
  `Vec<CellValue>`，保留 BeanMap 等纯值 API 的兼容边界
- [x] derive 宏为字段 Converter 生成独立的 `WriteCellData` 路径，不再在进入
  writer 前调用 `IntoExcelCell` 将样式和 `DataFormatData` 丢弃
- [x] Converter 前的 Handler 原值快照只对无副作用标量执行编码；URL、
  InputStream 等资源字段不再为制造快照而提前下载或消费，Converter 只执行一次
- [x] XLSX 普通写入将 Converter 样式作为目标样式，再按 Java 顺序合并注解和
  Handler 样式；运行期自定义格式串最终写入 `styles.xml`
- [x] BIFF8 普通写入将 Converter 的字体、填充和对齐请求交给 FONT/XF/Palette
  分配器；BIFF8 当前不支持任意自定义 FORMAT 记录，此限制保持明确
- [x] ZIP 模板保真写入保留与稀疏单元格对齐的 `WriteCellData` sidecar，
  编译并导入 Converter 产生的样式和自定义格式，不再只追加标量值
- [x] CSV 继续只输出标量文本，但通过 `effective_value` 保留 Converter 产生的
  公式、超链接、批注及图片基值语义；样式不写入无样式能力的 CSV 格式
- [x] `converter_write_cell_data_style_reaches_xlsx_xls_csv_and_template_backends`
  验证 derive 链仍保留绿色填充与 `0.0000`，普通 XLSX 和 ZIP 模板的
  `styles.xml` 均含对应记录，XLS/CSV 值可真实回读

本节修复的是已由 Java `AbstractExcelWriteExecutor.converterAndSet` 和
`FillStyleCellWriteHandler` 证明的生产语义差异，不是为了凑 API 增加空类型。
Converter、WriteCellData 及 Writer 总账本仍保持 `in_progress`，直到 Java
全部内建 Converter、BIFF8 自定义 FORMAT 和异常分支的 golden tests 完整对齐。

### 1.15 ExcelWriteFillExecutor 真实执行链（2026-07-24）

- [x] 移除仅保存 `WriteContext` 的标记式实现；Rust
  `ExcelWriteFillExecutor` 现在持有并调用实际 `WriteFillExecutor`，真实转发
  `fill` 与 `finish`，没有安装填充引擎时返回可见错误
- [x] `ExcelBuilderImpl::fill` 和资源收尾不再绕过这个 Java 对应类型，
  生产调用链统一经过 `ExcelWriteFillExecutor`
- [x] 模板分析、集合游标、相对行索引、横向/纵向填充、强制新行、自动样式和
  工作表 XML 修改继续由 `easyexcel-template::BuilderFillExecutor` /
  `ExcelTemplateWriter` 承担；writer 执行器作为真实适配层复用这些逻辑，
  避免跨 crate 循环依赖和重复实现
- [x] 执行器单测验证填充状态和 `finish(on_exception)` 均传入真实引擎，
  同时验证未安装引擎不是静默空操作
- [x] 门面测试
  `facade_do_fill_accepts_collection_config_and_supplier` 验证集合、配置和
  supplier 通过新执行链抵达模板引擎；模板执行器测试继续覆盖方向、
  `force_new_row` 和 `auto_style`

该项已经消除“同名空壳方法”问题，但账本仍保持 `in_progress`：Java
`ExcelWriteFillExecutor` 内部直接操作 POI 行列对象，而 Rust 将对应状态机拆分在
writer 适配层与 template 引擎中；完整升级仍需补齐复杂合并区域、跨工作表公式和
全部 Java 异常分支的 golden tests。

### 1.16 动态表头后的基础类型剩余列（2026-07-24）

- [x] 按 Java `ExcelWriteAddExecutor.addBasicTypeToExcel` 修复无模型行映射：
  先按有效 `headMap` 顺序消费数据，数据短于表头时停止创建后续单元格
- [x] 数据长于表头时，不再静默丢弃剩余值；从最大表头物理列的下一列继续写入，
  对齐 Java 为 issue #1702 增加的追加逻辑
- [x] `includeColumnIndexes + orderByIncludeColumn` 下，基础类型值按有效表头顺序
  连续消费，不再错误地用原始表头下标抽取数据
- [x] CSV writer 启用可变宽记录，允许“表头 N 列、数据 N+M 列”，避免公共映射
  修复后又被 CSV 固定宽度校验拒绝
- [x] `dynamic_basic_row_keeps_values_beyond_the_head_map_across_backends`
  真实验证 XLSX、BIFF8 XLS、CSV 和 ZIP 模板追加四条后端路径均保留表头外数据
- [x] 边界测试覆盖短于表头、等于表头、长于表头以及筛选后重排的列映射

这项修复的是生产数据丢失问题，不只是 `ExcelWriteAddExecutor` 的同名方法。
账本仍保持 `in_progress`，后续还需逐项对齐 Java `CollectionRowData` /
`MapRowData` 对非连续整数键、空行和 `null` 行的具体行为。

### 1.17 CollectionRowData / MapRowData 主写入闭环（2026-07-24）

- [x] 两个 Java 对应 RowData 类型实现真实 `ExcelRow`，不再只能作为
  `ExcelWriteAddExecutor` 私有辅助参数；可以直接传入静态写入和状态式
  `ExcelWriter::write`
- [x] `CollectionRowData` 保留集合顺序，支持读取侧 `RowData` 回建以及
  XLSX、XLS、CSV 写入
- [x] `MapRowData` 对齐 Java 的精确契约：以 `map.size()` 为行宽，再调用
  `map.get(0..size)`；非连续键不会被擅自解释成物理稀疏列
- [x] Rust 仍用 `DynamicRow` 提供明确的稀疏物理列能力，避免为了兼容 Java
  的 `MapRowData` 行为而丢失 Rust 原有实用功能
- [x] `ExcelWriteAddExecutor` 的 Map 分支改用相同的 `size + get(index)` 规则，
  主门面和命名执行器不再产生两套结果
- [x] `collection_and_map_row_data_enter_the_public_writer_backends` 验证：
  两种类型在同一状态式 XLSX 工作表分批追加、Collection 写 BIFF8 XLS、
  Map 写 CSV，以及非连续整数键的 Java 对齐结果

Collection/Map/RowData 账本仍为 `in_progress`，剩余主要缺口是 Java 集合中
`null` 整行的“保留相对行索引但不创建 Row/不触发 Handler”语义。

### 1.18 Java null 整行 / Rust Option 行语义（2026-07-24）

- [x] `ExcelRow` 增加行级缺席契约，`Option<T>` 作为 Java 集合元素
  `null` 的 Rust 表达；`Some(T)` 完整复用底层模型的 schema、metadata、
  Converter 和写入行为
- [x] `None` 仍推进物理行号及 `relativeRowIndex`，但不会调用
  `to_excel_write_row`、不会创建 Row/Cell，也不会触发 Row/Cell Handler
- [x] 普通 XLSX 和 BIFF8 XLS 均保留行号空洞；缺席行不应用行高、循环合并、
  gzip spill、样式或图片布局
- [x] CSV 不输出伪造的空记录，但内部物理行号与相对数据索引继续推进，使后续
  Handler 上下文与 Java 一致
- [x] ZIP/OOXML 模板和 BIFF8 模板均保留行号空洞；OOXML 不为缺席行生成空
  `<row>`，模板 Handler、行高和样式编译也跳过该位置
- [x] `absent_option_rows_keep_indexes_without_rows_cells_or_handlers` 验证
  XLSX、XLS、CSV 和 ZIP 模板的首行/缺席行/第三行序列，Handler 只观察到
  `(physical, relative) = (0,0), (2,2)`，并用会 panic 的模型证明缺席行未转换

该语义缺口已闭合；`ExcelWriteAddExecutor` 总账本仍为 `in_progress`，因为
JavaBean 反射字段补写顺序、全部 Converter 异常上下文以及表级 Head 合并仍需
更多 Java golden tests。

### 1.19 Java FieldCache 列顺序与写转换异常上下文（2026-07-24）

- [x] `ordered_columns` 改为复现
  `ClassUtils.buildSortedAllFieldMap`：显式 `@ExcelProperty.index` 先占用
  物理列，未指定 index 的字段按 `order`、声明顺序填入最小空闲列；不再把
  schema 声明位置错误地当成优先物理列
- [x] `java_field_cache_order_and_forced_index_are_preserved_across_backends`
  验证 order 相同字段的声明顺序、显式 index 空位跳过，以及 XLSX、BIFF8
  XLS、CSV、模板稀疏行的相同表头和数据布局
- [x] 派生宏把字段 Converter 的任意错误统一包装成位置化 `ExcelError::Data`；
  写后端再以实际 sheet、物理 row、筛选/重排后的 column 修正上下文，对齐
  `ExcelWriteDataConvertException(CellWriteHandlerContext, ...)`
- [x] 普通 XLSX、XLS、CSV 以及已有内容的 ZIP 模板均验证第二个失败数据行的
  sheet/row/column/field；模板错误行号以原模板最后一行之后开始计算
- [x] include/exclude 在派生模型转换前生效：被排除字段不再调用 Converter，
  对齐 Java 在 `FieldCache.sortedFieldMap` 阶段过滤字段的行为
- [x] `write_converter_errors_report_physical_sheet_row_column_and_field` 同时验证
  失败定位和 XLSX/XLS/CSV 被排除 Converter 不执行
- [x] `#[derive(ExcelRow)]` 在编译期拒绝两个字段声明相同的显式 index，
  对齐 Java `ClassUtils.declaredOneField` 初始化 `indexFieldMap` 时抛错

这两项 JavaBean 主写入语义已经进入真实生产后端。账本仍保持
`in_progress`：继承字段的 Rust 表达边界和 `ClassUtils` 兼容门面的空占位仍需
继续处理；手写 schema 的重复 index 与表级 Head 合并已在下一节闭合。

### 1.20 手写 schema 校验与 Java 多级表头合并（2026-07-24）

- [x] 手写 `ExcelRow::schema()` 与派生宏模型统一执行显式 index 唯一性校验；
  错误文本对齐 Java `ClassUtils.declaredOneField`，且发生在 include/exclude、
  Converter、模板解析、Handler 回调和文件创建之前
- [x] `duplicate_manual_indexes_fail_before_handlers_filters_templates_and_output`
  覆盖 XLSX、BIFF8 XLS、CSV、ZIP 模板和状态式 writer，并用会 panic 的
  `to_row()` 证明失败发生在行转换之前
- [x] `dynamic_head_merge_ranges` 逐语句迁移
  `ExcelWriteHeadProperty.headCellRangeList()` 的占用集合 + 贪心矩形算法；
  不再只做同层横向合并，支持横向、纵向及矩形合并
- [x] 动态表头短路径按 Java `ExcelHeadProperty.initHeadRowNumber()` 重复末级
  标题；算法保留 Java 的精确行为，包括不同父标题下同名末级仍可横向合并
- [x] include/exclude 与 `orderByIncludeColumn` 后，表头通过有效列的 source
  index 取值，不再要求原始 head 数量等于筛选后列数，也不再把 physical index
  错当原始表头下标
- [x] `automaticMergeHead(false)` 在 XLSX 与 XLS 路径真实关闭自动合并；
  `relativeHeadRowIndex` 同时移动表头单元格、合并范围、Handler 相对索引、
  样式/行高及后续数据行
- [x] ZIP OOXML 模板通过 `<mergeCells>` 写回自动合并，BIFF8 模板通过新增
  `MERGECELLS (0x00E5)` 记录保留并追加合并；两条路径均覆盖已有种子行和
  `relativeHeadRowIndex`
- [x] `dynamic_multi_level_head_merges_parents_and_offsets_data_rows` 真实读取
  XLSX/XLS 合并区域并验证关闭合并；`dynamic_head_merges_are_preserved_on_xlsx_and_xls_templates`
  真实读取两种模板输出并验证相同区域

`ExcelWriteHeadProperty` 账本仍保持 `in_progress`，因为 Rust 兼容属性类型本身
仍未承载 Java `Head` 的全部字段/注解元数据；本节只声明自动合并计算与主写入
后端语义已经闭合，不把“测试通过”误报为整个类型迁移完成。

### 1.21 ExcelWriteHeadProperty 真实属性模型（2026-07-24）

- [x] 删除原先只有三个松散字段的兼容占位实现；`ExcelWriteHeadProperty`
  现在通过 `Deref<ExcelHeadProperty>` 表达 Java 继承，真实持有
  `headClazz / headKind / headRowNumber / headMap`
- [x] `headRowHeightProperty`、`contentRowHeightProperty` 和
  `onceAbsoluteMergeProperty` 改为对应的强类型属性，不再错误地把整个
  `ExcelWriteMetadata` 塰入 `onceAbsoluteMerge`
- [x] `from_columns` 将派生宏产生的 `ExcelColumn` 解析为 Java `Head`：
  保存有效物理列、字段名、强制 index/name、字段级列宽/循环合并/表头样式/
  字体，并对列宽、样式和字体执行类级元数据回退
- [x] `headCellRangeList()` 的占用集合与贪心矩形算法迁入公开属性类型；
  XLSX、BIFF8 XLS 及两种模板后端不再维护私有平行算法，而是构造该属性并
  消费其 `CellRange`
- [x] 修正 `ExcelHeadProperty::for_class`：同时传入 class 和显式 head 时，
  Java 在字段初始化后仍为 `CLASS`，Rust 不再错误保留 `STRING`
- [x] 修正 `Head` 空字符串边界：Java 只拒绝 `null`，Rust `String` 无 null，
  因而允许空标题，不再把 empty 错当 null
- [x] `excel_write_head_property_resolves_metadata_and_java_merge_ranges` 验证
  类型化属性、字段/类级回退、短表头补齐及横向/纵向/矩形合并；
  原有真实 XLSX/XLS/模板回归继续通过，证明公开模型已进入生产路径

该类型不再是假方法或空壳。账本仍保持 `in_progress`，原因是 Java 构造器还会
通过反射把继承字段及注解对象装入 `Head.field`，Rust 当前以 derive schema
表达相同写入能力，但尚未完成“继承字段”对应的类型层迁移与 Java golden
constructor 测试；不能据此宣称整个 `ExcelWriteHeadProperty` 已 100% 完成。

### 1.22 WriteHolder 表头属性契约（2026-07-24）

- [x] `WriteHolder.excel_write_head_property()` 返回类型由错误的
  `ExcelWriteMetadata` 修正为真实 `ExcelWriteHeadProperty`
- [x] `AbstractWriteHolder` 不再保存 `Option<ExcelWriteMetadata>` 并在空值时
  返回静态伪属性；每个 holder 始终拥有一个可解析、可查询的表头属性对象
- [x] 增加显式 `resolve_head` / `set_excel_write_head_property`，使 builder 在
  schema、动态表头和类级元数据确定后可以安装完整属性，而非把元数据伪装成属性
- [x] 补齐 Java `WriteHolder` 的 `orderByIncludeColumn`、四个 include/exclude
  collection getter；接口调用者不再只能依赖 `ignore()` 猜测 holder 配置
- [x] `holder_exposes_real_head_property_and_complete_selection_surface` 通过 trait
  object 验证真实 class 多级表头、强类型行高和全部列选择集合

本节闭合的是接口/存储类型错误。账本仍为 `in_progress`：当前生产后端直接以
`WriteOptions + ExcelRow schema` 构建有效列，尚未把 workbook/sheet/table 三层
live holder 全部统一为该兼容 holder；后续需迁移 Java `AbstractWriteHolder`
构造器完整父子覆盖顺序，并用状态式多 sheet/table 上下文测试证明接线完成。

### 1.23 Workbook / Sheet / Table Holder 继承链（2026-07-24）

- [x] 三个具体 holder 均真实内嵌 `AbstractWriteHolder`，并通过
  `Deref/DerefMut` 暴露继承属性；不再只是名称、编号和游标组成的平行空壳
- [x] 新增各层 `from_parameter`：workbook 解析根默认值，sheet 继承 workbook，
  table 继承 sheet；缺省值继承、显式值覆盖、显式空集合清除父集合
- [x] 各层保留独立可变 holder 状态，表级后续可安装自己的
  `ExcelWriteHeadProperty`，不会修改父 sheet/workbook
- [x] `workbook_sheet_table_holders_resolve_java_parent_chain` 验证
  `needHead` 的根值/覆盖、include indexes 的继承/清空、exclude names 的继承
  及 `orderByIncludeColumn` 表级覆盖
- [x] 既有 table builder 5 项测试继续通过，证明嵌入继承状态未破坏公开构建 API

账本继续保持 `in_progress`：这些 owned mirror holder 已具有真实继承语义，但
`ExcelWriter` 的运行时资源仍由内部 sheet state 持有；下一步要把初始化后的
mirror holder 与运行时 handler context 使用同一份解析状态，并验证切换
sheet/table 时 `currentWriteHolder()` 的动态类型与 Java 一致。

### 1.24 Live currentWriteHolder 统一状态（2026-07-24）

- [x] 将纯元数据 `ExcelWriteHeadProperty` 下沉到 `easyexcel-core`，消除
  `WriteContextHolder` 暴露真实表头属性时的 core → writer 循环依赖；
  `easyexcel_writer::ExcelWriteHeadProperty` 原导入路径继续 re-export，保持兼容
- [x] `WriteContextHolder` 不再只有 path/sheet/table 浅视图，现可查询
  `Holder` 类型、真实 `ExcelWriteHeadProperty`、need/automatic merge/relative
  head row，以及全部 include/exclude/order 状态
- [x] `WriteContextImpl` 持有 `WriteContextHolderState`；sheet/table 切换会同步
  `Holder::Sheet` / `Holder::Table`，不再靠 `table_no.is_some()` 由调用者猜测
- [x] `ExcelBuilderImpl::write_rows<T>` 在每次实际写入前复用生产路径的
  `selected_columns` 和动态表头 source-index 规则，解析筛选/重排后的
  `HeadMap`、字段名、类型元数据和类级行高，再更新同一个 live context
- [x] table 同时继承 sheet 的 include indexes 并叠加自己的 include field
  names，保留 Java 四类集合分别继承、最终按 OR 选择字段的精确行为
- [x] template fill 也调用相同 holder 更新函数；write 和 fill 不再产生两套
  `currentWriteHolder()` 状态
- [x] `live_current_write_holder_tracks_resolved_sheet_and_table_state` 使用同一个
  builder 从 sheet 切换到 table，验证真实输出文件、holder 类型、有效列顺序、
  动态多级表头、类型行高及父子 include 合并
- [x] `template_fill_updates_the_same_live_current_holder` 验证 fill 的 sheet 名称
  trim、needHead、automaticMergeHead、relativeHeadRowIndex 与动态表头

`WriteContext / WriteContextImpl / ExcelBuilderImpl` 账本仍保持 `in_progress`：
Java 上下文还暴露 workbook/sheet/table holder 的具体对象、POI workbook/sheet
和 deprecated 访问器；Rust 已闭合业务可观察的当前 holder 配置，但尚未对齐全部
具体对象访问面和异常完成路径。

### 1.25 Handler callback 共享 currentWriteHolder 状态（2026-07-24）

- [x] `WriteHolderContext` 现在实现 `WriteContext` 和 `WriteContextHolder`；
  workbook/sheet/row/cell handler 可直接使用
  `context.write_context().current_write_holder()`，不再只能读取路径、sheet 名称和行列
- [x] callback holder 快照携带同一套 `WriteContextHolderState`，可查询
  `Holder` 动态层级、`ExcelWriteHeadProperty`、need/automatic merge/relative head row
  以及 include/exclude/order 配置
- [x] 抽取 `resolved_write_context_holder_state<T>`，`ExcelBuilderImpl` 的公开
  `write_context()` 与 XLSX/XLS/CSV/template handler 构造路径复用同一解析算法，
  避免两条路径对动态表头 source index、字段顺序和类型元数据产生漂移
- [x] stateful writer 的 sheet/table scope 和一次性 XLSX、BIFF8、CSV、模板写入
  全部在创建 row/cell callback 前解析有效 holder 状态
- [x] `table_holder_runs_supplementary_callbacks_then_own_parent_row_chain` 通过真实 table
  写入，在 workbook、row、cell 回调内验证 `Holder::Workbook/Holder::Table`、输出路径、
  table 编号、needHead、automaticMergeHead、include 顺序、exclude 字段及过滤后的
  `HeadMap`

这里使用不可变、逐回调快照，而不是伪造 Apache POI 对象或跨 callback 暴露可变
writer 借用；它保证 handler 在该生命周期点观察到与 Java 一致的已解析配置。
账本仍保持 `in_progress`：Java context 中的 POI `Workbook/Sheet/Row/Cell` 具体类型
由 Rust 后端中立 handle 代替，`WriteHandlerUtils.create*Context(WriteContext, ...)`
兼容入口也尚未全部改造成由 live context 构造。

### 1.26 WriteHandlerUtils live context 工厂（2026-07-24）

- [x] 对齐 Java 四个同名工厂的首参数：`createWorkbook/Sheet/Row/CellWriteHandlerContext`
  现在接收 `&dyn WriteContext`，从 `currentWriteHolder()` 克隆 workbook/sheet/table
  holder 图，而不再把 path 或 sheet name 伪装成完整上下文
- [x] 原 Rust 简化构造入口以
  `*_from_path`、`*_from_name`、`*_from_sheet` 后缀保留，兼容显式创建轻量测试上下文
- [x] `WriteContextHolderState::from_holder` 统一复制 head、need/merge/relative、
  include/exclude/order 状态；`WriteHolderContext::from_write_context` 同步 path、
  sheet name/no、last row、table no 和动态 holder 类型
- [x] sheet/row/cell 工厂在尚未选择 sheet 时返回明确 `ExcelError::Format`，不制造
  空 sheet holder；与 Java 仅在合法写入生命周期调用这些方法的前置条件一致
- [x] 实际 `HandlerHolderScope` 改为先构建 live `WriteHolderContext`，再调用
  `with_write_context` 建立 sheet/row/cell callback，生产路径和公开工厂共享同一快照语义
- [x] `java_style_context_factories_clone_the_live_holder_graph` 验证完整
  workbook→sheet→table 图、表头和列配置；`java_style_sheet_row_and_cell_factories_reject_missing_sheet`
  验证非法生命周期；真实 table handler 回归继续通过

仍保持 `in_progress`：Java `WriteHandlerUtils` 还从 `AbstractWriteHolder` 选择 own/effective
handler execution chain；Rust 的 chain 选择目前由 `ExcelWriter` 的
`effective_sheet_handlers/effective_table_handlers` 完成，尚未收敛到 core 工具层。

### 1.27 AbstractWriteHolder handler execution chain（2026-07-24）

- [x] 新增统一 `HandlerExecutionScope`，显式保存 Java `AbstractWriteHolder` 的
  `own` 与 `effective` 两组执行链；workbook/sheet 的补充回调选择 own，
  row/cell 以及正常 holder 生命周期选择 effective
- [x] 父子链严格按 `table own → sheet own → workbook own/default` 合并，使用稳定
  `order()` 排序；同 order 时子 holder 先于父 holder
- [x] `NotRepeatExecutor.uniqueValue()` 在每条链构造时真实去重，不再通过散落调用点
  临时替换成 noop；重复值由排序后首个、即最具体 holder 的 handler 获胜
- [x] stateful workbook/sheet/table 的创建、写入与 finish 都从统一 execution scope
  选择 handler；删除 `effective_sheet_handlers/effective_table_handlers` 平行入口
- [x] 修正首次写入前 `register_write_handler` / `prepend_write_handlers` 未同步
  `current_effective_handlers` 的边界：空 workbook 直接 `finish()` 时，新注册 handler
  现在同时收到 create 和 dispose 回调
- [x] `holder_handler_scope_deduplicates_effective_chain_but_runs_each_own_chain` 验证
  workbook/sheet/table 使用相同 unique value 时，各层 own 补充回调均执行，而
  sheet effective、row effective、finish effective 只执行最具体层
- [x] `table_holder_runs_supplementary_callbacks_then_own_parent_row_chain` 继续验证
  非重复、同 order handler 的 table→sheet→workbook 稳定顺序；
  `handler_registered_before_empty_finish_participates_in_dispose_chain` 覆盖空写完成路径

这里的执行链已进入 XLSX/XLS/CSV stateful 实际写入路径，不是只存在于
`handler/chain/*.rs` 的兼容类型。账本继续保持 `in_progress`：Java
`AbstractWriteHolder` 还包含注解 handler 装载、POI holder 资源和样式缓存等职责；
四个独立 chain 兼容类型本身也尚未取代后端中立的统一 `WriteHandler` 分发接口。

### 1.28 四类 HandlerExecutionChain 完整生命周期（2026-07-24）

- [x] `WorkbookHandlerExecutionChain` 从仅有 dispose 补齐
  `beforeWorkbookCreate / afterWorkbookCreate / afterWorkbookDispose`
- [x] `SheetHandlerExecutionChain` 从仅有 after 补齐
  `beforeSheetCreate / afterSheetCreate`
- [x] `RowHandlerExecutionChain` 从仅有 dispose 补齐
  `beforeRowCreate / afterRowCreate / afterRowDispose`
- [x] `CellHandlerExecutionChain` 从仅有 before 补齐
  `beforeCellCreate / afterCellCreate / afterCellDataConverted / afterCellDispose`
- [x] 原实现错误调用 `before_cell / after_row / after_sheet / after_workbook`
  兼容 hook；现全部调用与 Java 同名的精确生命周期方法，保留 trait 默认委托兼容
- [x] 新增 `with_handler` 对应 Java 的带 handler 构造器，同时保留 Rust 原有空
  `new()`；`add_last` 仍维护注册顺序
- [x] `all_java_chain_lifecycle_methods_forward_in_registration_order` 用两级链逐一验证
  12 个 Java 生命周期方法及 first→next 的执行顺序

四个 chain 类型不再是只有一个方法的占位结构。它们仍保持 `in_progress`，因为
生产 stateful writer 使用上一节的后端中立 `HandlerExecutionScope` 执行同一语义，
尚未让公开 linked-list chain 直接持有共享 handler 实例；强行改用独占
`Box<dyn WriteHandler>` 会破坏 Java 父子 holder 共享同一 handler 对象的行为。

### 1.29 AbstractWriteHolder 注解 Handler 装载（2026-07-24）

- [x] 新增 `load_annotation_handlers<T>`，按 Java `initAnnotationConfig` 的顺序从
  `ExcelRow::schema/write_metadata` 真实构造 loop merge、column width、annotation
  cell style、row height、once absolute merge 五类 handler，不再只依赖后端散落读取元数据
- [x] stateful `ExcelWriter` 分别持久化 sheet/table 注解 handler，并将其置于用户
  custom handler 之前；own/effective scope 统一进行稳定排序和父子 holder 合并
- [x] 对齐 Java order：annotation style/width/height 为 `-60000/-50000`，
  `LoopMergeStrategy` 修正为默认 `0`（原 Rust 错设为 `50000`）；同 order 下子级
  注解 handler 先于父级自定义 handler，因此后者可覆盖默认列宽、表头和内容行高
- [x] `WriteHandler::style_loop_merge` 暴露可执行循环合并配置；XLSX、BIFF8 和模板
  后端统一收集 options、注解 fallback、handler 三路配置并按属性去重，避免同一
  `A2:A3` 或绝对区域被重复写入
- [x] `annotation_config_loads_real_ordered_handlers_for_every_java_strategy` 验证五类
  handler 及 order；`stateful_sheet_persists_annotation_handlers_and_deduplicates_merges`
  验证持久化和最终 OOXML 唯一区域；`parent_custom_dimension_handler_overrides_annotation_defaults`
  验证 workbook 父级 custom handler 的覆盖行为
- [x] XLSX template、XLS template、BIFF8 样式/合并及注解尺寸回归继续通过

账本继续保持 `in_progress`：一次性公开写函数仍接收借用的
`&mut [Box<dyn WriteHandler>]`，无法在不改变兼容 API/所有权的前提下把新建注解
handler 注入调用方切片，因此该路径保留等价的直接元数据 fallback；stateful
sheet/table 已使用真实共享 handler。Java POI 样式缓存和具体资源对象也尚未迁移。

### 1.30 Holder converter map 与默认写转换器（2026-07-24）

- [x] `WriteContextHolder` 和 `WriteContextHolderState` 新增真实
  `converter_map()`，workbook/sheet/table 当前 holder 均向 handler 和 executor
  暴露最终生效的转换器注册表，对齐 Java
  `ConfigurationHolder.converterMap()` / `WriteContext.currentWriteHolder()`
- [x] `AbstractWriteHolder` 根 holder 从 `load_default_write_converter()` 初始化；
  子 holder 克隆父级注册表后再合并自身 converter，保持
  `table → sheet → workbook → default` 的最近注册优先覆盖语义
- [x] stateful writer 的实际数据转换和 live holder context 使用同一组 workbook、
  sheet、table converter；不再出现 handler 查到的 converter 与最终写入所用
  converter 不一致
- [x] `DefaultConverterLoader.loadDefaultWriteConverter()` 不再返回空注册表：
  BigDecimal、BigInt、布尔、Rust 有符号/无符号整数、浮点、日期时间、字符串、
  PathBuf、字节数组和 URL 均注册真实 `IntoExcelCell` 转换器
- [x] `load_default_read_converter()` 不再返回空注册表：按
  `(Rust 目标类型, CellDataType)` 注册 BigDecimal、BigInt、布尔、整数、浮点、
  日期时间和 String 的真实 `FromExcelCell` 转换器；`ReadOptions::default()` 直接
  装载该表，自定义 converter 后注册并覆盖默认项
- [x] 数值目标对 Boolean 单元格补齐 Java 语义：`false/true` 分别转换为 `0/1`，
  覆盖普通整数、浮点、BigInt 和 BigDecimal
- [x] `ReadOptions.use_1904_windowing` 沿 `RowData → ConvertContext` 进入默认日期
  converter；数字单元格现可转换为 NaiveDate/NaiveDateTime，并复现 Excel 1900
  虚构闰日边界及 1904 epoch。真实 XLSX 数字 `0/1.5` 回读验证为
  `1904-01-01` 和 `1904-01-02 12:00:00`
- [x] `#[excel(use_1904_windowing = true/false)]` 将字段级日期窗口写入
  `ExcelColumn`，并在 `RowData → ConvertContext` 中优先于全局
  `ReadOptions.use_1904_windowing`。真实 XLSX 用例在同一行验证字段级 `true`
  得到 `1904-01-01`、字段级 `false` 得到 `1900-01-01 12:00:00`
- [x] `ConverterRegistry` 写入键扩展为 Java 的
  `(Rust TypeId, target CellDataType)`；无目标类型 converter 继续作为自定义 fallback，
  后注册覆盖规则不变
- [x] CSV 写入在字段 converter 执行前选择 `CellDataType::String`，数字、布尔、
  日期时间使用专用 String converter；XLS/XLSX 保持无目标类型的原生单元格转换。
  `csv_uses_java_target_string_converter_before_cell_handlers` 验证 handler 看到的已是
  String，而非 CSV 序列化阶段才临时格式化
- [x] stateless `write_xlsx*` / `write_xls*` 文件与流入口统一物化“默认 converter
  + 调用方 converter”注册表，不再只把自定义表传给 `ExcelRow`；真实 XLSX/XLS
  回读用例同时验证默认注册表已经进入行转换路径
- [x] sheet 写入后继续 table 写入时，父子注解生成的同一
  `OnceAbsoluteMergeStrategy` 区域只应用一次；XLSX、BIFF8 和 template 路径共享
  跨 holder 去重逻辑，修复 `A11:B11` 重叠合并错误
- [x] `default_write_registry_contains_real_scalar_converters` 验证默认 i32/String
  转换器真实执行；`converter_map_clones_parent_and_applies_child_override` 验证父级
  克隆与子级覆盖；`live_holder_converter_map_matches_sheet_and_table_write_precedence`
  同时探测 live holder 并用 Calamine 回读实际文件；既有 facade converter 优先级
  用例继续通过

本项仍保持 `in_progress`。Java 使用独立的 `@DateTimeFormat` 注解承载字段级
格式与 1904 windowing，Rust 按语言习惯将其合并到 `#[excel(...)]`；两者的字段
覆盖语义已经对齐，但注解对象并非直接同型。Java 还有 `java.util.Date` 与
InputStream 等 Rust 不存在直接同型对象，需继续按 Rust 对等类型核对。

### 1.31 日期转换器逐类实装（2026-07-24）

- [x] `LocalDateDateConverter`、`LocalDateNumberConverter`、
  `LocalDateStringConverter` 不再是零字段占位类型，均实现真实
  `Converter<NaiveDate>`；Date 写入携带默认 `yyyy-MM-dd` 数据格式，Number
  双向处理 1900/1904 serial，String 支持字段格式和 Java `switchDateFormat`
  常用回退格式
- [x] `LocalDateTimeDateConverter`、`LocalDateTimeNumberConverter`、
  `LocalDateTimeStringConverter` 实现真实 `Converter<NaiveDateTime>`，并已替换
  默认注册表中的匿名泛型日期转换器
- [x] `DateDateConverter`、`DateNumberConverter`、`DateStringConverter` 使用
  独立 `JavaDate` newtype 对齐 `java.util.Date`；它与 `NaiveDateTime`
  (`LocalDateTime`) 拥有不同 `TypeId`，两组转换器可以同时进入默认注册表
- [x] BIFF8 日期写序号修正为 POI 边界：1900-01-01 = 1、
  1900-02-28 = 59、1900-03-01 = 61、1904-01-01 = 0；此前固定使用
  1899-12-30 epoch 会把 1900 年 1–2 月写大一天
- [x] `number_converters_are_real_bidirectional_java_equivalents`、
  `string_converters_honor_field_format_and_reject_invalid_input`、
  `date_cell_converters_attach_java_default_data_formats` 直接实例化九个镜像类，
  覆盖 Number/String/Date、格式、非法输入和 1900/1904 语义

### 1.32 Boolean 转换器逐类实装（2026-07-24）

- [x] `BigDecimalBooleanConverter`、`BigIntegerBooleanConverter`、
  `ByteBooleanConverter`、`ShortBooleanConverter`、`IntegerBooleanConverter`、
  `LongBooleanConverter`、`FloatBooleanConverter`、`DoubleBooleanConverter` 均由空壳
  改为真实双向 `Converter<T>`：Boolean 读取为 1/0，写入时仅精确等于 1 为
  `true`，2、-1、0 均为 `false`
- [x] `BooleanBooleanConverter`、`BooleanNumberConverter`、
  `BooleanStringConverter` 分别实现 Boolean/Number/String source key。Number
  严格复现 Java `BigDecimal.ONE.compareTo(...) == 0`；String 严格复现
  `Boolean.valueOf`，因此 `"true"` 忽略大小写为 true，而 `"1"`、带空格文本和
  任意其他文本均为 false
- [x] `StringBooleanConverter` 实现 Boolean→`"true"/"false"` 以及 String→Boolean，
  写入同样采用 Java `Boolean.valueOf` 规则
- [x] 默认读注册表已替换为上述具体转换器；默认 Boolean 写和 CSV String 写分别
  使用 `BooleanBooleanConverter`、`BooleanStringConverter`，不再依赖语义更宽松的
  通用 `FromExcelCell<bool>`
- [x] 三个直接调用测试覆盖全部 12 个镜像类型，并在默认注册表测试中固定
  `number 2 → false`、`string "1" → false` 的 Java 语义

### 1.33 Number 转换器逐类实装（2026-07-24）

- [x] `BigDecimalNumberConverter`、`BigIntegerNumberConverter`、
  `ByteNumberConverter`、`ShortNumberConverter`、`IntegerNumberConverter`、
  `LongNumberConverter`、`FloatNumberConverter`、`DoubleNumberConverter` 均由空壳
  改为真实双向 `Converter<T>`，并接入默认读写注册表
- [x] 数值读取严格以 Java `ReadCellData.getNumberValue()` 的 `BigDecimal` 语义为
  基准；整数目标先向零截断，再复现 `byteValue`、`shortValue`、`intValue`、
  `longValue` 的低位二进制回绕，而不是采用 Rust `TryFrom` 的范围错误
- [x] 数值写入统一生成 `CellValue::Decimal`，对应 Java
  `NumberUtils.formatToCellData` 的 `new BigDecimal(num.toString())`；字段存在
  number format 时同步写入 `WriteCellData.dataFormatData`
- [x] 保留 easyexcel-rust 的无损增强：转换器仍返回 Java 对应的 Decimal，
  但 XLSX、BIFF8 与模板后端发现整数超出 Excel `2^53-1` 精确范围时改写为文本，
  避免 `i64::MAX/MIN` 和任意精度 `BigInt` 在文件往返时静默失真
- [x] 非 NUMBER source 会返回带 sheet/row/column/field 的定位错误，NaN 和正负
  Infinity 写入会像 Java `new BigDecimal(...)` 一样失败
- [x] 直接调用测试覆盖 8 个镜像类型、255→Byte -1、2^32-1→Integer -1、
  2^64-1→Long -1、负数回绕、任意精度写入、格式元数据和默认注册优先级

### 1.34 数值 String 转换器与 DecimalFormat（2026-07-24）

- [x] `BigDecimalStringConverter`、`BigIntegerStringConverter`、
  `ByteStringConverter`、`ShortStringConverter`、`IntegerStringConverter`、
  `LongStringConverter`、`FloatStringConverter`、`DoubleStringConverter` 已由空壳
  改为真实双向转换器，并替换默认注册表中的通用 String 转换
- [x] 无格式读取严格使用完整 `BigDecimal` 输入，不再 `trim()`；因此
  `"1.00"` 可按 Java 向零截断为整数，`" 1.00"`、`"1.00 "` 和尾随非数字会像
  `new BigDecimal(string)` 一样失败
- [x] 格式化与解析层覆盖 Java `DecimalFormat` 的必需语义：`0/#`、可选/必需
  小数、分组、百分比、千分比、科学计数、引号字面量、正负子模式、前后缀以及
  `ParsePosition` 风格的前缀解析
- [x] 新增 `NumberRoundingMode`，完整覆盖 Java `UP`、`DOWN`、`CEILING`、
  `FLOOR`、`HALF_UP`、`HALF_DOWN`、`HALF_EVEN`、`UNNECESSARY`；负数
  CEILING/FLOOR 的方向已按 Java 校验
- [x] `#[excel(number_format = "#.##%", rounding_mode = "HALF_UP")]` 会生成静态
  `ExcelColumn` 元数据并进入读写 Converter 上下文；原有 `format = ...` 仍兼容
- [x] Float/Double 无格式写入复现 Java `toString` 的 `.0`、科学计数阈值、
  `-0.0`、`NaN`、`Infinity`；有格式 Infinity 使用 Java DecimalFormat 的 `∞`
  和格式前后缀
- [x] 本地 Java 运行黄金值固定 `#.##%`、`#`、`0.00`、`#,##0.00`、
  `0.00;[neg]0.00`、`0.00E00` 以及宽松解析边界；Core 直接测试覆盖全部 8 个
  镜像类型和默认注册优先级

### 1.35 String Number/String/Error 转换器（2026-07-24）

- [x] `StringNumberConverter` 已从零字段占位类型改为真实双向
  `Converter<String>`：NUMBER 读取遵循 Java 的显式日期格式、显式数字格式、
  单元格自带 Excel 显示格式、默认 BigDecimal 文本优先级；String 写 NUMBER
  使用严格 `BigDecimal` 解析，不接受首尾空白
- [x] `ReadConverterContext` 新增源单元格 `display_value` 与精确
  `decimal_value`，derive 生成的模型映射会从 `RowData` 传递这两项元数据；
  百分比、日期等 POI/DataFormatter 显示不再在进入 converter 前丢失，也不会因
  `f64` 中转破坏 OOXML 十进制文本
- [x] `StringStringConverter` 与 `StringErrorConverter` 已实现严格 source-type
  读取和对应 String/Error 写入；默认读写注册表改用三个具体转换器
- [x] XLSX `t="e"` 与 Calamine Error 不再提前降成普通 String，而是保留
  `CellValue::Error`，确保 `(String, ERROR)` 注册键能够真实命中
- [x] 端到端 XLSX 模型测试固定 `0.1250 + 0.00% → "12.50%"` 与
  `#DIV/0! → String`，覆盖 reader、格式元数据、derive、默认注册表和 converter
  完整链路

### 1.36 图片转换器逐类实装（2026-07-24）

- [x] `ByteArrayImageConverter` 与 `BoxingByteArrayImageConverter` 已实现真实
  `Converter<Vec<u8>>` / `Converter<Box<[u8]>>`，转换结果使用
  `WriteCellData::from_image`，对应 Java `WriteCellData(byte[])` 的图片列表语义，
  不再只是普通 `CellValue::Image` 门面
- [x] `FileImageConverter` 读取 `PathBuf` 指向的完整文件并传播 I/O 错误；
  `StringImageConverter` 保留 Java String path 语义
- [x] Java 包路径下的 `InputStreamImageConverter`、`UrlImageConverter` 已重导出
  真实实现；流转换消费当前 reader 剩余字节但不关闭调用方对象，URL 保持 Java
  默认 connect 1 秒/read 5 秒超时并确保响应 reader 生命周期结束
- [x] 默认写注册表已使用 byte array、boxed byte array、File、URL 的具体图片
  converter；直接注册测试固定 `WriteCellData.value == Empty` 且 image list 保存原始
  字节，端到端 XLSX 测试覆盖 byte array、file、input stream 与 URL 路径

### 1.37 Converter 基础设施与 nullable 分派（2026-07-24）

- [x] `AutoConverter` 现在真实实现泛型 `Converter<T>` 的默认 unsupported 行为，
  对应 Java 注解默认值“只触发按类型自动查找、不可直接调用”的 sentinel 语义
- [x] `ConverterKeyBuild` 从 `type_name` 字符串拼接占位改为可哈希、可比较的
  `ConverterKey(TypeId, Option<CellDataType>)`；读注册按完整二元键选择，写注册同时
  支持 Java 的无目标类型键与指定目标单元格类型键
- [x] `NullableObjectConverter<T>` 从空结构体改为真实 marker trait；注册表新增
  `register_nullable` / `register_nullable_for_write_type`，普通 converter 遇到
  EMPTY 或 `Option::None` 会被跳过，nullable converter 才会收到空值
- [x] facade 和 sheet builder 暴露 `register_nullable_converter`；derive 对
  `Option<T>` 字段将 null 状态传入注册表，保持 Java `AbstractExcelWriteExecutor`
  的空值门控顺序
- [x] `ReadConverterContext`、`WriteConverterContext` 已进入真实分派链，分别携带
  单元格/显示值/精确十进制/公式/字段/分析上下文和写入值/字段/写上下文
- [x] Core 直接测试固定普通与 nullable 转换器在空读、空写上的分歧；真实 XLSX
  端到端测试证明普通 converter 不会制造空值内容，而 nullable converter 将
  `None` 写成 `"nullable-null"` 并可被模型回读

### 1.38 java.util.Date 独立注册键（2026-07-24）

- [x] 新增透明 `JavaDate(NaiveDateTime)` 对等类型及双向转换，避免把 Java
  `java.util.Date` 与 `java.time.LocalDateTime` 都压缩成同一个 Rust `TypeId`
- [x] 三个 `Date*Converter` 改为 `Converter<JavaDate>`，继续复用已验证的
  1900/1904 serial、格式切换与默认日期时间数据格式逻辑
- [x] 默认写注册表同时注册 `JavaDate → DateDateConverter` 和
  `NaiveDateTime → LocalDateTimeDateConverter`；目标 String 写注册和
  Number/String 读注册也保持两个独立键
- [x] 默认注册表测试对同一日期值同时查询两个类型键；真实 XLSX 模型包含
  `JavaDate` 与 `NaiveDateTime` 两列并完整往返，证明不存在后注册覆盖冲突

### 1.39 InputStream 默认图片转换链（2026-07-24）

- [x] `ImageInputStream` 增加默认的 `Box<dyn Read + Send>` 类型擦除形态及
  `boxed(reader)` 构造器，使 Java 字段类型 `InputStream` 在 Rust 中拥有稳定
  `TypeId`，可进入 `DefaultConverterLoader` 的默认写注册表
- [x] `InputStreamImageConverter` 生成 `WriteCellData::from_image`，不再把图片
  字节仅包装成缺少 image list 的普通 `CellValue::Image`
- [x] 流的首次转换会消费并缓存剩余字节，derive 的原始值回调与 converter 写入
  两个阶段复用同一结果，不会二次读取后产生空图片；底层 reader 最终位置仍保持
  EOF，且库不会替调用方关闭 reader
- [x] `FromExcelCell` 明确返回 unsupported，保持 Java
  `InputStreamImageConverter` 只写不读的能力边界，而不是提供虚假的读取方法
- [x] 无显式 `converter` 注解的真实 XLSX 测试验证默认注册表自动创建
  `xl/media/*` 与 drawing anchor，并逐字节核对嵌入媒体内容

### 1.40 CSV/XLS/XLSX 格式化读取上下文（2026-07-24）

- [x] `CsvReadContext`、`XlsReadContext`、`XlsxReadContext` 均暴露共享
  `AnalysisContextImpl`、格式专用 workbook holder 与当前 sheet holder，保持
  Java interface 的三组窄化 getter
- [x] 三个 `Default*ReadContext` 构造时从 `ReadOptions` 真实传播 charset、
  auto-close、ignore-empty-row 与 password，不再创建丢失配置的默认 holder
- [x] 三种 context 的 `current_sheet` 同时更新共享分析上下文与格式专用 sheet
  holder；sheet 编号、名称和 listener context 保持一致
- [x] 重复选择同一 sheet 统一返回 Java 文案
  `Cannot read sheet repeatedly.`，不是静默覆盖或始终返回 `None`
- [x] 统一契约测试覆盖 CSV/XLS/XLSX 的格式类型、配置传播、初始空 holder、
  sheet 切换和重复读取错误

### 1.41 Empty 双包标记类型（2026-07-24）

- [x] Java `com.alibaba.excel.Empty` 与 `com.alibaba.excel.support.Empty` 都是无字段
  标记类；Rust 使用同一个零尺寸 `Empty`，support 包路径仅做类型重导出
- [x] 直接测试验证 root/support 两个路径拥有相同 `TypeId`、零字节大小及默认构造、
  Copy、相等语义，不以两个无关占位结构伪造包兼容

### 1.42 EasyExcel 继承入口（2026-07-24）

- [x] Java `EasyExcel` 本身没有新增方法，仅继承 `EasyExcelFactory`；Rust 将
  `EasyExcelFactory` 定义为 `EasyExcel` 的同类型别名，避免复制两套可能漂移的
  静态 facade
- [x] 契约测试验证两个入口 `TypeId` 相同，并可通过 `EasyExcelFactory` 调用真实
  `writer_table` 与 `writer_sheet_index` 构建逻辑
- [x] `EasyExcelFactory.java` 的 26 个 Java overload 已逐项映射到 Rust 的
  无歧义命名入口；`Class head` 由 `T: ExcelRow` 泛型表达，listener 由
  `register_read_listener` 或既有 `read::<T, L>` 表达

### 1.43 Ignore、日期与数字格式注解（2026-07-24）

- [x] `ExcelIgnore` / `ExcelIgnoreUnannotated` 保持 Java 无字段 marker 语义，同时
  提供可构造、可比较的零尺寸 Rust 元数据类型；derive 分别映射为字段
  `#[excel(ignore)]` 与类型 `#[excel(ignore_unannotated)]`
- [x] `DateTimeFormat` 不再是空 marker，真实保存 Java `value=""` 与
  `BooleanEnum::Default`，并支持显式格式和 1904 windowing 三态
- [x] `NumberFormat` 不再是空 marker，真实保存 Java `value=""` 与默认
  `HALF_UP`，覆盖全部 `NumberRoundingMode`（含 `UNNECESSARY`）
- [x] derive 将 number/date pattern、rounding mode、1904 override 写入
  `ExcelColumn`，被忽略字段与未注解字段不会进入 schema/read/write 映射
- [x] 真实 XLSX 往返测试同时验证 ignore、ignore-unannotated、BigDecimal 数字格式
  和日期格式字段；被忽略字段回读为 Rust `Default`，不是隐藏列或空占位列

### 1.44 EasyExcelFactory 全 overload 与输入/输出流（2026-07-24）

- [x] 新增 `reader()`、`reader_from_path()`、`reader_from_input_stream()` 和
  `writer()`、`writer_to_path()`、`writer_to_output_stream()`，底层复用真实
  `ExcelReaderBuilder` / `ExcelWriterBuilder`，不是只保存参数的空门面
- [x] 补齐 `readSheet()`、`writerSheet()` 的默认、序号、名称、序号+名称四种
  Rust 命名映射，以及 `writerTable()` 默认/序号 builder
- [x] Java `InputStream` 映射会读取到自动清理的临时文件，并由 `ExcelReader`
  持有 `TempPath` 守卫；这使现有需要 seek 的 XLSX/XLS 解析器可直接复用，
  `finish/drop` 后删除临时输入
- [x] 输入流通过 ZIP/OLE 魔数分流；OLE 中存在 `EncryptedPackage` 与
  `EncryptionInfo` 时识别为加密 XLSX，而不是误判成 BIFF `.xls`
- [x] 契约测试执行真实路径写入、内存输出流写入、普通/加密 XLSX 输入流读取、
  listener 回调、sheet/table 构建以及临时文件清理

### 1.45 ExcelReader 与格式 executor 生命周期（2026-07-24）

- [x] `ExcelReader` 覆盖 Java `read()`（deprecated alias）、`readAll()`、
  `read(ReadSheet...)` / `read(List<ReadSheet>)`（Rust slice）、`analysisContext()`、
  `getAnalysisContext()`、`excelExecutor()`、`finish()`、`close()` 与 `Drop`
- [x] `ExcelAnalyserImpl` 不再只保存 `ExcelTypeEnum` 后直接调用自由函数；它实际
  持有 `ExcelReadExecutorKind::{Xlsx,Xls,Csv}`，sheet discovery 与 typed
  listener 解析通过同一个格式 executor
- [x] XLSX executor 返回真实工作簿 sheet 清单，CSV executor 返回实际单逻辑
  sheet；`close()` 幂等，关闭后再次读取返回明确错误
- [x] `read(&[])` 保持 Java 空 sheet 参数错误，指定多个 sheet 时按工作簿顺序
  匹配并应用 sheet 级 head/scientific 参数

### 1.46 ExcelAnalyser 接口去占位（2026-07-24）

- [x] 删除原先无参数 `analysis()` 只记录 `Unsupported` 的占位行为；Rust trait
  现在以泛型参数显式接收 typed listener，并调用保存的真实格式 executor
- [x] Java 把 listener 擦除后放在 `ReadWorkbook`，Rust 以
  `analysis::<T, L>(&mut listener)` 表达同一职责；sheet/read-all 选择仍由调用前
  的 `ReadOptions`/`ExcelReader` 设置
- [x] `ExcelReader::read_all`、附加 listener 和逐 sheet 路径统一经过
  `ExcelAnalyser` trait，不再绕开接口直接调用实现细节

### 1.47 ExcelReadExecutor 与 CSV executor 去占位（2026-07-24）

- [x] `ExcelReadExecutor::execute()` 改为
  `execute::<T, L>(options, listener) -> Result<()>`；这是 Java 从
  `ReadWorkbook` 获取擦除 listener/options 的 Rust 泛型化表达
- [x] `ExcelReadExecutorKind::{Xlsx,Xls,Csv}` 的 execute 统一分派到对应真实
  OOXML、BIFF、CSV 解析器，不再由 analyser 重复按类型调用自由函数
- [x] `XlsxSaxAnalyser` / `XlsSaxAnalyser` 原来只记录 Unsupported 的 void
  execute 已删除；trait 与兼容 `execute_with_listener` 均运行同一解析逻辑
- [x] `CsvExcelReadExecutor` 持有实际路径并复用 charset/BOM、sheet 选择、空值、
  trim 和 listener 生命周期所在的公共 CSV reader；未绑定路径返回配置错误

### 1.48 ExcelAnalyserImpl 真实清理生命周期（2026-07-24）

- [x] `finish()` 不再只设置 `finished`：会清理数字 formatter 缓存和磁盘共享
  字符串缓存的当前线程文件句柄，并立即释放 Java `InputStream` 物化临时文件
- [x] 临时输入守卫下沉到 analyser，因此解析失败触发 Java 式自动 `finish()`
  时也能立即删除文件，而不必等待整个 `ExcelReader` 被 drop
- [x] 磁盘缓存 TLS 句柄记录其文件路径；切换缓存文件时重新打开，避免连续读取
  不同工作簿却错误复用上一份缓存文件
- [x] Rust 解析器的 ZIP/CSV/工作簿句柄均为单次 `execute` 局部所有权并由 RAII
  关闭；密码也只保存在单 reader options 中，不存在 POI
  `Biff8EncryptionKey` 的全局/TLS 密码状态
- [x] analyser 通过透明 listener 转发器保留每次 head、row、extra、exception、
  hasNext 与完成回调的最新 `AnalysisContext`，`analysisContext()` 不再永远返回
  构造时的空上下文
- [x] 修正对象矩阵中不存在的 `remove_thread_local_cache()` /
  `clear_encrypt_03()` 方法虚报，改为记录实际 Rust 语义适配

### 1.49 XlsListSheetListener 去空壳（2026-07-24）

- [x] `XlsListSheetListener` 从空 unit struct 改为持有
  `DefaultXlsReadContext`、路径、options 与扫描结果的真实 metadata listener
- [x] 构造时按 Java 设置 `needReadSheet = false`，`execute()` 实际枚举 XLS
  工作表并写入 `ReadWorkbookHolder.actualSheetDataList`
- [x] `XlsSaxAnalyser::new` 统一通过该 listener 完成预扫描，不再绕过对应对象
- [x] `XlsSaxAnalyser.processRecord` 已不再返回 `Unsupported`，并由真实
  OLE `/Workbook`/`/Book` 物理记录流按 SID 分派到 19 个 handler；未知 SID
  依 Java 语义忽略，损坏的 header/payload 会返回格式错误
- [x] `IgnorableXlsRecordHandler` 已由可实例化空 struct 改为 marker trait；
  工作表 BOF 按 `First/Index/Name/All` 更新跳过状态，未选中 sheet 的 17 类
  ignorable record 不再执行；`AbstractXlsRecordHandler` 同样改为不可实例化 trait
- [ ] `XlsSaxAnalyser` 仍保持 `in_progress`：BIFF handler 已进入主执行链并产生
  可观察状态，但 typed row 最终仍由 calamine 物化；SST/LabelSST 与
  Formula/String 的跨记录解析已闭合，但尚未写入 live XLS context，公式 token
  也尚未还原为文本；Dummy missing-cell/row、Obj/TxO/Note/Hyperlink 等协作及
  XLS extra listener 仍待迁移

### 1.50 XLS 真实 BIFF record dispatch（2026-07-24）

- [x] 新增受边界校验的 BIFF record walker，复用现有 OLE Workbook stream，
  不再另造一套文件打开逻辑；`xls_display` 同时复用该底层输入
- [x] `execute()` 在 typed XLS 行读取前实际完成物理 record dispatch，真实
  Java 多 sheet fixture 验证 workbook/worksheet BOF、BoundSheet 名称、EOF
  与 handler 命中，不只依赖手写 payload
- [x] Number/Blank/BoolErr/RK handler 保存真实解析结果，Number 同时保留 XF
  index；BoundSheet handler 完整解析 BIFF8 compressed/UTF-16 sheet name
- [x] `XlsRecordHandler.process_record` 改为必实现方法，具体 handler 无法再继承
  默认空实现；Java 的 Dummy record 因为是 POI 合成事件，保留显式专用入口
- [x] SST + Continue 与 Formula + String cached result 的物理记录协作已在
  1.51 节完成
- [ ] 下一批优先完成 MissingRecordAware dummy 事件以及 Obj/TxO/Note 的
  shape/text 关联，再把 record 输出写入 live XLS context

### 1.51 XLS SST 与 Formula 跨记录状态（2026-07-24）

- [x] 实现 BIFF8 `XLUnicodeRichExtendedString` 分段解码：支持 compressed/
  UTF-16、CONTINUE 中切换字符宽度、rich text run 和 ExtRst 跨记录跳过
- [x] SST 不再只保存 unique count；dispatcher 汇聚物理 CONTINUE 后构造有序
  shared-string cache，损坏或未结束的逻辑记录在下一记录/EOF 返回格式错误
- [x] LabelSST 解析 row/column/SST index，通过上述 cache 解析实际字符串并
  应用 Java `autoTrim`，缺失 index 保持 Empty 语义
- [x] Formula 解析 numeric/string/boolean/error/empty 五类 cached result、
  row/column/XF index；String record 支持 CONTINUE，并完成前一条 string
  formula 的 pending cell
- [x] 真实 Java multisheet XLS fixture 验证 SST 非空且 decoded 数量等于
  unique count；边界测试覆盖 continuation width 切换、rich/ext、截断和
  Formula→String/LabelSST 协作
- [ ] Formula token (`rgce`) 到公式文本目前仍由 calamine 主读取链提供，Rust
  record handler 尚无 `HSSFFormulaParser` 等价实现；上述 record cell 也尚未
  成为 listener 的唯一数据源，因此四个具体 handler 继续保持 `in_progress`

## 2. 待继续（不得删减既有实现）

- [ ] 将 `in_progress` 账本项补齐 `test_evidence` 后升为 `complete`
- [ ] 逐步把残留 `mod.rs` 迁移为 `foo.rs + foo/`（66 处，渐进）
- [ ] Java `easyexcel-test` 全量用例 → Rust parity/golden（Phase E）
- [ ] `migration-audit-strict` 全绿（Phase G / v1.0）

### 2.1 基线外兼容类型

`ExcelReaderTableBuilder` / `ReadTable` 不存在于当前 Java 4.0.3
基线源码中，因此 Rust 侧保留实现属于旧版本兼容扩展，不能作为
4.0.3 的 1:1 迁移完成证据；相关 Phase E 测试注释后续需改为
`legacy compatibility`，且不会加入 325 文件主账本。

## 3. 验收命令

```bash
cargo run -p xtask -- migration-audit
cargo check --workspace --all-targets
cargo test -p easyexcel-reader
cargo test -p easyexcel-reader builder::excel_reader_sheet_builder::tests
cargo test -p easyexcel-writer write::builder::excel_writer_builder::tests
cargo test -p easyexcel-writer builder::excel_writer_table_builder::tests
cargo test -p easyexcel-writer excel_builder::tests::fill_config_initializes_java_defaults_and_preserves_overrides
cargo test -p easyexcel-template builder_fill_executor::tests::builder_fill_config_propagates_direction_force_row_and_auto_style
cargo test -p easyexcel excel_builder::tests::facade_do_fill_accepts_collection_config_and_supplier
cargo test -p easyexcel reader_builder_register_read_listener_dispatches_in_registration_order
cargo test -p easyexcel --test codegraph_phaseE_metadata_1to1_tests read_sheet_test
cargo test -p easyexcel-core excel_download_error_body -- --nocapture
cargo test -p easyexcel-web-axum -- --nocapture
cargo test -p easyexcel-web-actix -- --nocapture
```
