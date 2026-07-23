# EasyExcel ↔ easyexcel-rust 测试对比迁移状态

> 更新时间：2026-07-20（P0-3 DONE：core STRING 产物饱和；golden **110**）

## 结论（诚实）

**仍未达到严格 100%（含 Ehcache stress、xls 写细节、large07 全量、Fill composite / FillAnnotation 样式断言）。**  
本轮将 Java golden **103 → 110**（门槛 ≥108），补齐仍缺格式变体与 Fill complex。  
**P0-3（core 缺类 STRING 读写产物）判定 DONE / 饱和**——见下方清单。  
无本机 JDK 时 `cargo test -p easyexcel --test java_golden_tests` 仍绿（缺文件硬失败，禁止 soft-skip）。  
新增用例均为**全表** `rows` 对照；`ofNoRows = 0`。

## P0-3 DONE：core 类 STRING 覆盖清单

| core 包 | Golden 产物 | 饱和说明 |
|---------|-------------|---------|
| cache | `cache_data` / `_xls` / `_csv` | STRING 读写已齐；invoke/MEMORY 反射改注解属行为测，不进 golden |
| celldata | `celldata_data` / `_xls` / `_csv` | 全表（含 CN DateTimeFormat） |
| charset | `charset_gbk` / `charset_utf8` | CSV 编码全表 |
| exception | `exception_data{,_xls,_csv}` + `exception_stop_sheet0..4` | 内容 + 多 sheet；抛错路径属行为测 |
| handler | `handler_data` / `_xls` / `_csv` | STRING 内容；回调副作用属行为测 |
| large | `large_sample` / `_xls` / `_csv` | 100×25；**large07 74MB 故意不纳入** |
| nomodel | `nomodel_data{,_xls,_csv}` + `nomodel_repeat{,_xls,_csv}` | 含 repeat 写回三格式 |
| noncamel | `noncamel_data` / `_xls` / `_csv` | 已齐 |
| parameter | `parameter_data` / `_csv` / `_xls` | 已齐 |
| repetition | `repetition_data{,_xls,_csv}` + `repetition_table` | 已齐 |
| skip | `skip_sheet0..3` | 四 sheet 全表；csv 多 sheet 写 Java 即抛错 |

**仍非 STRING-golden 目标（不算 P0-3 未饱和）：** `ConverterTest`（无文件）、`FillAnnotation*` / `FillStyle*`（POI 样式/合并/图片断言）、`FillDataTest#composite`（FillWrapper 复合）、`large07`、cache 反射改注解、exception 抛错、handler 回调。

## Java golden 对照表（`tests/golden/*.expected.json`）

| # | Golden | Java source / 场景 |
|---|--------|-------------------|
| 1–8 | `compatibility_t01_xls` … `t09` | CompatibilityTest |
| 9–10 | `bom_*` | BomDataTest |
| 11–15 | `demo_*` | demo / extra / cellData / simple07 |
| 16–19 | `simple_*` | SimpleDataTest |
| 20–25 | `converter_*` | ConverterDataTest fixture + write 三格式 |
| 26–29 | `multiplesheets_*` | MultipleSheetsDataTest |
| 30–34 | `dataformat_*` | DateFormatTest（全表）+ date1/date2 |
| 35–36 | `template_*` | TemplateDataTest |
| 37–45 | `fill_simple` / `_xls` / `fill_horizontal` / `_xls` / `fill_by_name` / `_xls` / `fill_complex` / `_xls` | FillDataTest |
| 46–47 | `style_data` / `_xls` | StyleDataTest |
| 48 | `annotation_data` | AnnotationDataTest |
| 49–53 | `exclude_*` / `include_*` | ExcludeOrIncludeDataTest |
| 54–56 | `no_head_data` / `_xls` / `_csv` | NoHeadDataTest |
| 57 | `sort_data` | SortDataTest |
| 58 | `encrypt_data` | EncryptDataTest |
| 59–61 | `cache_data` / `_xls` / `_csv` | CacheDataTest |
| 62–64 | `celldata_data` / `_xls` / `_csv` | CellDataDataTest |
| 65–66 | `charset_*` | CharsetDataTest |
| 67–74 | `exception_*` | ExceptionDataTest |
| 75–77 | `handler_data` / `_xls` / `_csv` | WriteHandlerTest |
| 78–80 | `large_sample` / `_xls` / `_csv` | LargeDataTest 小样本 |
| 81–86 | `nomodel_*` | NoModelDataTest（含 repeat 三格式） |
| 87–89 | `noncamel_*` | UnCamelDataTest |
| 90–92 | `parameter_*` | ParameterDataTest |
| 93–95 | `complex_head_*` | ComplexHeadDataTest |
| 96–98 | `annotation_index_name_*` | AnnotationIndexAndNameDataTest |
| 99–101 | `list_head_*` | ListHeadDataTest |
| 102–105 | `repetition_*` | RepetitionDataTest |
| 106–109 | `skip_sheet0..3` | SkipDataTest |
| — | （合计 **110**） | `scripts/export-java-golden.sh` |

刷新：`./scripts/export-java-golden.sh`（需 JDK + Maven）。提交 `expected.json` + `artifacts/` 后，无 JDK 亦可跑 Rust 侧对照。

## 本轮新增（相对 103，+7）

- `cache_data_xls` / `cache_data_csv`
- `nomodel_repeat_xls` / `nomodel_repeat_csv`
- `fill_by_name_xls`
- `fill_complex` / `fill_complex_xls`（headRowNumber=3，forceNewRow + LoopMerge）

## 集成测试规模（约）

| 套件 | 用例 |
|------|------|
| java_golden | **110** expected.json + 专项断言（门槛 ≥108） |
| bom / demo / demo_write_extra | 3 + 25 + 12 |
| temp_1to1_tests | **134 passed / 1 ignored**（`CacheTest#cache` JVM PersistentCacheManager） |

## 距 100% 剩余（非 P0-3）

1. **temp 仍 ignore（1）**：`CacheTest#cache`（`org.ehcache.PersistentCacheManager` 探针，非 EasyExcel `Ehcache` 门面）。`HeadReadTest#testCache` 已改为 `ReadCacheMode::Disk` 三次 XLS 读断言。
2. **xls 写**能力仍在收敛 —— Java 写 `.xls` 以 Rust **读**对照为主。
3. **large07**（74MB）未纳入 golden（体积/耗时）。
4. **Fill composite / FillAnnotation / FillStyle** —— 复合 FillWrapper、样式/合并/图片 POI 断言；非纯 STRING 产物。
5. **cache invoke/MEMORY、exception 抛错、handler 回调** —— 行为测已在其它套件。

## 验证

```bash
./scripts/export-java-golden.sh   # 可选：有 JDK 时刷新
cargo test -p easyexcel --test java_golden_tests
# expected: all passed；*.expected.json ≥ 108；ofNoRows = 0
```
