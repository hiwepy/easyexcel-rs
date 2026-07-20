# temp 包方法级 1:1 命名矩阵

> Java：`com.alibaba.easyexcel.test.temp` 全部 `*Test.java` 的每个 `@Test`
> Rust：`crates/easyexcel/tests/temp_1to1_tests/`

## 汇总

| 指标 | 数量 |
|------|------|
| temp `@Test` 总数 | 134 |
| implemented | 132 |
| ignored | 2 |
| missing（命名缺口） | **0** |

## 分类摘要（本轮）

| 分类 | 说明 | 数量 |
|------|------|------|
| **可移植已激活** | 去 `#[ignore]` + fixture/tempfile/真断言 | 132（含本轮 +30：本机路径→fixture / 可契约化） |
| **Ehcache exclusion（仍 ignore）** | CacheTest / HeadReadTest#testCache；严格 100% 下排除（见 test-parity-status） | 2 |
| **纯 POI（仍 ignore）** | 无（本轮已把可替换路径与契约项全部激活） | 0 |

默认：`cargo test -p easyexcel --test temp_1to1_tests` → **132 passed / 2 ignored / 0 failed**。

## 本轮解锁清单（≥30，相对本轮前 32 ignored）

### 本机路径 → `tests/fixtures/`（禁止 `/Users/...` / `D:\...`）

| Rust fn | Fixture / 替代 |
|---------|----------------|
| `poi_poi_2_test_test` / `_last_row_num_xssf` | `fill/simple.xlsx` |
| `poi_poi_format_test_last_row_num` | `fill/simple.xlsx` |
| `poi_poi_format_test_last_row_num_xssf` | `dataformat/dataformat.xlsx` |
| `poi_poi_test_last_row_num` / `_xssf` / `_233` / `_cp` / `_233443` / `_2333` / `_testread` / `_2332222` / `_23443` / `_2` / `_xssf_2` | `fill/simple.xlsx` |
| `poi_poi_test_last_row_num_xss_fv_22` | `xls/converter03.xls` |
| `poi_poi_test_last_row_num_255` | `fill/complex.xlsx` |
| `poi_poi_test_testread_read` | `fill/simple.xlsx` bytes |
| `poi_poi_write_test_part_4` | `converter/img.jpg`（原 HTTP URL） |
| `write_temp_write_test_image_write_poi` / `_tep` | `converter/img.jpg` |

### 可契约化（EasyExcel API / 纯逻辑 smoke）

| Rust fn | 契约 |
|---------|------|
| `poi_poi_3_test_encryption` / `_encryption_2` | password encrypt round-trip |
| `poi_poi_write_test_write_0` / `_write` | 大整数 **字符串** 写读（避 f64 精度） |
| `poi_poi_write_test_write_01` | float→decimal smoke |
| `poi_poi_write_test_write_1` | long2bytes |
| `poi_poi_write_test_part` / `_part_2` | `${...}` fill placeholder |
| `large_temp_large_data_test_t_04_write_excel_poi` | EasyExcel large write |

### 仍 ignore（Ehcache — 严格 100% 标准下 **exclusion**）

> 定策见 `docs/test-parity-status.md`「temp Ehcache：严格 100% 用户标准下的策略表」。  
> **不**用 `ReadCacheMode` smoke 冒充 Ehcache；测试层保留同名 `#[ignore]`。

| Java FQCN#method | Rust fn | 排除理由 |
|------------------|---------|----------|
| `com.alibaba.easyexcel.test.temp.cache.CacheTest#cache` | `cache_cache_test_cache` | 纯 Ehcache PersistentCacheManager 10GB disk；无 EasyExcel API |
| `com.alibaba.easyexcel.test.temp.read.HeadReadTest#testCache` | `read_head_read_test_test_cache` | `readCache(new Ehcache(20))`；Rust 无 Ehcache 等价实现 |

## 矩阵

| Java FQCN#method | Rust fn | Status | Reason |
|-----------------|---------|--------|--------|
| `com.alibaba.easyexcel.test.temp.FillTempTest#complexFill` | `fill_temp_test_complex_fill` | implemented | fill complexFill |
| `com.alibaba.easyexcel.test.temp.FillTempTest#complexFillWithTable` | `fill_temp_test_complex_fill_with_table` | implemented | fill complexFillWithTable |
| `com.alibaba.easyexcel.test.temp.Lock2Test#test` | `lock_2_test_test` | implemented | converter07 fixture read |
| `com.alibaba.easyexcel.test.temp.Lock2Test#test33` | `lock_2_test_test_33` | implemented | EasyExcel style/write portable |
| `com.alibaba.easyexcel.test.temp.Lock2Test#write` | `lock_2_test_write` | implemented | handler style write |
| `com.alibaba.easyexcel.test.temp.Lock2Test#simpleWrite` | `lock_2_test_simple_write` | implemented | handler style write |
| `com.alibaba.easyexcel.test.temp.Lock2Test#testc` | `lock_2_test_testc` | implemented | CellReference via position_utils |
| `com.alibaba.easyexcel.test.temp.Lock2Test#simpleRead` | `lock_2_test_simple_read` | implemented | demo.xlsx fixture read |
| `com.alibaba.easyexcel.test.temp.Lock2Test#test2` | `lock_2_test_test_2` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.Lock2Test#test335` | `lock_2_test_test_335` | implemented | A1 cell ref parse portable |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt` | `lock_2_test_numberforamt` | implemented | excel serial date probes |
| `com.alibaba.easyexcel.test.temp.Lock2Test#testDate` | `lock_2_test_test_date` | implemented | date epoch smoke |
| `com.alibaba.easyexcel.test.temp.Lock2Test#testDateAll` | `lock_2_test_test_date_all` | implemented | sampled excel date roundtrip |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt3` | `lock_2_test_numberforamt_3` | implemented | dataformat date fixtures |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt4` | `lock_2_test_numberforamt_4` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt77` | `lock_2_test_numberforamt_77` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt99` | `lock_2_test_numberforamt_99` | implemented | datetime nanos format smoke |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt5` | `lock_2_test_numberforamt_5` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt6` | `lock_2_test_numberforamt_6` | implemented | decimal scale smoke |
| `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt7` | `lock_2_test_numberforamt_7` | implemented | decimal scale smoke |
| `com.alibaba.easyexcel.test.temp.LockTest#test` | `lock_test_test` | implemented | demo.xlsx fixture read |
| `com.alibaba.easyexcel.test.temp.LockTest#test2` | `lock_test_test_2` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.StyleTest#test` | `style_test_test` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.StyleTest#poi` | `style_test_poi` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.StyleTest#poi07` | `style_test_poi_07` | implemented | dataformat.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.StyleTest#poi0701` | `style_test_poi_0701` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.StyleTest#poi0702` | `style_test_poi_0702` | implemented | dataformat.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.StyleTest#poi0703` | `style_test_poi_0703` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.StyleTest#testFormatter` | `style_test_test_formatter` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.StyleTest#testFormatter2` | `style_test_test_formatter_2` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.StyleTest#testFormatter3` | `style_test_test_formatter_3` | implemented | date/number format write-read |
| `com.alibaba.easyexcel.test.temp.StyleTest#testBuiltinFormats` | `style_test_test_builtin_formats` | implemented | builtin format style write |
| `com.alibaba.easyexcel.test.temp.WriteLargeTest#test` | `write_large_test_test` | implemented | large write/read |
| `com.alibaba.easyexcel.test.temp.WriteLargeTest#read` | `write_large_test_read` | implemented | large07 fixture present |
| `com.alibaba.easyexcel.test.temp.WriteLargeTest#read2` | `write_large_test_read_2` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.WriteLargeTest#read3` | `write_large_test_read_3` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.WriteLargeTest#test2` | `write_large_test_test_2` | implemented | large batched write |
| `com.alibaba.easyexcel.test.temp.WriteV33Test#handlerStyleWrite` | `write_v_33_test_handler_style_write` | implemented | handler style write |
| `com.alibaba.easyexcel.test.temp.WriteV33Test#test4` | `write_v_33_test_test_4` | implemented | handler style write |
| `com.alibaba.easyexcel.test.temp.WriteV34Test#test` | `write_v_34_test_test` | implemented | handler style write |
| `com.alibaba.easyexcel.test.temp.Xls03Test#test` | `xls_03_test_test` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.Xls03Test#test2` | `xls_03_test_test_2` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.cache.CacheTest#cache` | `cache_cache_test_cache` | ignored | **exclusion**: Ehcache PersistentCacheManager 10GB（见 test-parity-status） |
| `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#write` | `csv_csv_read_test_write` | implemented | csv write read |
| `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#read1` | `csv_csv_read_test_read_1` | implemented | csv fixture read |
| `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#csvWrite` | `csv_csv_read_test_csv_write` | implemented | csv write read |
| `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#writev2` | `csv_csv_read_test_writev_2` | implemented | csv write read |
| `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#writeFile` | `csv_csv_read_test_write_file` | implemented | CSV fixture FileMagic probe |
| `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#read` | `csv_csv_read_test_read` | implemented | csv fixture read |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test` | `dataformat_data_format_test_test` | implemented | dataformat xlsx |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#testxls` | `dataformat_data_format_test_testxls` | implemented | dataformat xls |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test3` | `dataformat_data_format_test_test_3` | implemented | dataformat.xlsx |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test31` | `dataformat_data_format_test_test_31` | implemented | is_a_date_format probes |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test43` | `dataformat_data_format_test_test_43` | implemented | locale date format smoke |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test463` | `dataformat_data_format_test_test_463` | implemented | locale date format smoke |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test1` | `dataformat_data_format_test_test_1` | implemented | is_a_date_format probes |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test2` | `dataformat_data_format_test_test_2` | implemented | vec clear vs realloc smoke |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test355` | `dataformat_data_format_test_test_355` | implemented | dataformat.xlsx |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#test3556` | `dataformat_data_format_test_test_3556` | implemented | dataformat.xlsx |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#tests` | `dataformat_data_format_test_tests` | implemented | locale date format smoke |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#tests1` | `dataformat_data_format_test_tests_1` | implemented | dataformat date fixtures |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#tests3` | `dataformat_data_format_test_tests_3` | implemented | locale date format smoke |
| `com.alibaba.easyexcel.test.temp.dataformat.DataFormatTest#tests34` | `dataformat_data_format_test_tests_34` | implemented | locale date format smoke |
| `com.alibaba.easyexcel.test.temp.fill.FillTempTest#simpleFill` | `fill_fill_temp_test_simple_fill` | implemented | fill simple |
| `com.alibaba.easyexcel.test.temp.fill.FillTempTest#listFill` | `fill_fill_temp_test_list_fill` | implemented | fill list |
| `com.alibaba.easyexcel.test.temp.fill.FillTempTest#complexFill` | `fill_fill_temp_test_complex_fill` | implemented | fill complexFill |
| `com.alibaba.easyexcel.test.temp.fill.FillTempTest#complexFillWithTable` | `fill_fill_temp_test_complex_fill_with_table` | implemented | fill complexFillWithTable |
| `com.alibaba.easyexcel.test.temp.fill.FillTempTest#horizontalFill` | `fill_fill_temp_test_horizontal_fill` | implemented | fill horizontal |
| `com.alibaba.easyexcel.test.temp.fill.FillTempTest#compositeFill` | `fill_fill_temp_test_composite_fill` | implemented | fill composite |
| `com.alibaba.easyexcel.test.temp.issue1662.Issue1662Test#test1662` | `issue_1662_issue_1662_test_test_1662` | implemented | issue1662 |
| `com.alibaba.easyexcel.test.temp.issue1663.FillTest#TestFillNullPoint` | `issue_1663_fill_test_test_fill_null_point` | implemented | fill issue1663 |
| `com.alibaba.easyexcel.test.temp.issue2443.Issue2443Test#IssueTest1` | `issue_2443_issue_2443_test_issue_test_1` | implemented | issue2443 |
| `com.alibaba.easyexcel.test.temp.issue2443.Issue2443Test#IssueTest2` | `issue_2443_issue_2443_test_issue_test_2` | implemented | issue2443 |
| `com.alibaba.easyexcel.test.temp.issue2443.Issue2443Test#parseIntegerTest1` | `issue_2443_issue_2443_test_parse_integer_test_1` | implemented | issue2443 parse |
| `com.alibaba.easyexcel.test.temp.issue2443.Issue2443Test#parseIntegerTest2` | `issue_2443_issue_2443_test_parse_integer_test_2` | implemented | issue2443 parse |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#read` | `large_temp_large_data_test_read` | implemented | large07 fixture present |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#noModelRead` | `large_temp_large_data_test_no_model_read` | implemented | no-model dynamic read |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#noModelRead2` | `large_temp_large_data_test_no_model_read_2` | implemented | no-model dynamic read |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04Write` | `large_temp_large_data_test_t_04_write` | implemented | large write/read |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04WriteExcel` | `large_temp_large_data_test_t_04_write_excel` | implemented | large write/read |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04WriteExcelNo` | `large_temp_large_data_test_t_04_write_excel_no` | implemented | large write/read |
| `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04WriteExcelPoi` | `large_temp_large_data_test_t_04_write_excel_poi` | implemented | large write/read contract |
| `com.alibaba.easyexcel.test.temp.poi.Poi2Test#test` | `poi_poi_2_test_test` | implemented | fill/simple.xlsx fixture (was D:\珠海.xlsx) |
| `com.alibaba.easyexcel.test.temp.poi.Poi2Test#lastRowNumXSSF` | `poi_poi_2_test_last_row_num_xssf` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.Poi3Test#Encryption` | `poi_poi_3_test_encryption` | implemented | EasyExcel password encrypt contract |
| `com.alibaba.easyexcel.test.temp.poi.Poi3Test#Encryption2` | `poi_poi_3_test_encryption_2` | implemented | EasyExcel password encrypt contract |
| `com.alibaba.easyexcel.test.temp.poi.PoiDateFormatTest#read` | `poi_poi_date_format_test_read` | implemented | dataformat date fixtures |
| `com.alibaba.easyexcel.test.temp.poi.PoiEncryptTest#encrypt` | `poi_poi_encrypt_test_encrypt` | implemented | encrypt |
| `com.alibaba.easyexcel.test.temp.poi.PoiEncryptTest#encryptExcel` | `poi_poi_encrypt_test_encrypt_excel` | implemented | encrypt |
| `com.alibaba.easyexcel.test.temp.poi.PoiFormatTest#lastRowNum` | `poi_poi_format_test_last_row_num` | implemented | fill/simple.xlsx fixture (was D:\原文件.xlsx) |
| `com.alibaba.easyexcel.test.temp.poi.PoiFormatTest#lastRowNumXSSF` | `poi_poi_format_test_last_row_num_xssf` | implemented | dataformat.xlsx fixture (was Downloads path) |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum` | `poi_poi_test_last_row_num` | implemented | fill/simple.xlsx fixture (was /Users/.../test3.xlsx) |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNumXSSF` | `poi_poi_test_last_row_num_xssf` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNumXSSFv22` | `poi_poi_test_last_row_num_xss_fv_22` | implemented | xls fixture read |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum233` | `poi_poi_test_last_row_num_233` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum255` | `poi_poi_test_last_row_num_255` | implemented | fill/complex.xlsx fixture (was D:\complex.xlsx) |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#cp` | `poi_poi_test_cp` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum233443` | `poi_poi_test_last_row_num_233443` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum2333` | `poi_poi_test_last_row_num_2333` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#testread` | `poi_poi_test_testread` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#testreadRead` | `poi_poi_test_testread_read` | implemented | fill/simple.xlsx bytes |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum2332222` | `poi_poi_test_last_row_num_2332222` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum23443` | `poi_poi_test_last_row_num_23443` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum2` | `poi_poi_test_last_row_num_2` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNumXSSF2` | `poi_poi_test_last_row_num_xssf_2` | implemented | fill/simple.xlsx fixture |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write0` | `poi_poi_write_test_write_0` | implemented | large number write/read |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write01` | `poi_poi_write_test_write_01` | implemented | float/decimal smoke |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write` | `poi_poi_write_test_write` | implemented | large number write/read |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write1` | `poi_poi_write_test_write_1` | implemented | long2bytes smoke |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#part` | `poi_poi_write_test_part` | implemented | fill placeholder pattern |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#part2` | `poi_poi_write_test_part_2` | implemented | fill placeholder pattern |
| `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#part4` | `poi_poi_write_test_part_4` | implemented | converter/img.jpg fixture (was HTTP URL) |
| `com.alibaba.easyexcel.test.temp.read.CommentTest#comment` | `read_comment_test_comment` | implemented | demo.xlsx fixture read |
| `com.alibaba.easyexcel.test.temp.read.HeadReadTest#test` | `read_head_read_test_test` | implemented | demo.xlsx fixture read |
| `com.alibaba.easyexcel.test.temp.read.HeadReadTest#testCache` | `read_head_read_test_test_cache` | ignored | **exclusion**: EasyExcel.readCache(Ehcache) 无 Rust 对等（见 test-parity-status） |
| `com.alibaba.easyexcel.test.temp.simple.HgTest#hh` | `simple_hg_test_hh` | implemented | repeat fixture |
| `com.alibaba.easyexcel.test.temp.simple.HgTest#hh5` | `simple_hg_test_hh_5` | implemented | repeat write |
| `com.alibaba.easyexcel.test.temp.simple.HgTest#hh2` | `simple_hg_test_hh_2` | implemented | repeat fixture |
| `com.alibaba.easyexcel.test.temp.simple.RepeatTest#hh` | `simple_repeat_test_hh` | implemented | repeat fixture |
| `com.alibaba.easyexcel.test.temp.simple.RepeatTest#hh2` | `simple_repeat_test_hh_2` | implemented | repeat fixture |
| `com.alibaba.easyexcel.test.temp.simple.RepeatTest#hh1` | `simple_repeat_test_hh_1` | implemented | repeat fixture |
| `com.alibaba.easyexcel.test.temp.simple.Write#simpleWrite` | `simple_write_simple_write` | implemented | relativeHeadRowIndex write/read |
| `com.alibaba.easyexcel.test.temp.simple.Write#simpleWrite1` | `simple_write_simple_write_1` | implemented | BeanMap field-key contract (str23/str22) |
| `com.alibaba.easyexcel.test.temp.simple.Write#simpleWrite2` | `simple_write_simple_write_2` | implemented | handler write; protectSheet Unsupported |
| `com.alibaba.easyexcel.test.temp.simple.Write#simpleWrite3` | `simple_write_simple_write_3` | implemented | dynamic head + cell handler; POI style Unsupported |
| `com.alibaba.easyexcel.test.temp.simple.Write#json` | `simple_write_json` | implemented | serde_json field-case serialize |
| `com.alibaba.easyexcel.test.temp.simple.Write#json3` | `simple_write_json_3` | implemented | serde_json field-case parse |
| `com.alibaba.easyexcel.test.temp.simple.Write#tableWrite` | `simple_write_table_write` | implemented | WriteTable 3-arg Unsupported + sheet write/read |
| `com.alibaba.easyexcel.test.temp.write.TempWriteTest#write` | `write_temp_write_test_write` | implemented | write newline |
| `com.alibaba.easyexcel.test.temp.write.TempWriteTest#cglib` | `write_temp_write_test_cglib` | implemented | write newline |
| `com.alibaba.easyexcel.test.temp.write.TempWriteTest#imageWrite` | `write_temp_write_test_image_write` | implemented | image write |
| `com.alibaba.easyexcel.test.temp.write.TempWriteTest#imageWritePoi` | `write_temp_write_test_image_write_poi` | implemented | converter/img.jpg + write smoke |
| `com.alibaba.easyexcel.test.temp.write.TempWriteTest#tep` | `write_temp_write_test_tep` | implemented | converter/img.jpg + write smoke |
| `com.alibaba.easyexcel.test.temp.write.TempWriteTest#large` | `write_temp_write_test_large` | implemented | large write/read |

## 验证

```bash
cargo test -p easyexcel --test temp_1to1_tests
cargo test -p easyexcel --test temp_1to1_tests -- --include-ignored --list
```
