//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.poi.*`
//!
//! Local-path / contractable probes use repo fixtures + EasyExcel APIs.
//! Pure Apache POI / Ehcache stress remains `#[ignore]` elsewhere.

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.poi.Poi2Test#test`
/// (was `D:\test\珠海.xlsx` — now `fill/simple.xlsx` fixture)
#[test]
fn poi_poi_2_test_test() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.Poi2Test#lastRowNumXSSF`
#[test]
fn poi_poi_2_test_last_row_num_xssf() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.Poi3Test#Encryption`
/// (POI agile encrypt of large07 → EasyExcel password round-trip)
#[test]
fn poi_poi_3_test_encryption() {
    helpers::assert_large_fixture();
    helpers::assert_encrypt();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.Poi3Test#Encryption2`
/// (BIFF8 password open → EasyExcel password round-trip)
#[test]
fn poi_poi_3_test_encryption_2() {
    helpers::assert_encrypt();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiDateFormatTest#read`
#[test]
fn poi_poi_date_format_test_read() {
    helpers::assert_dataformat_dates();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiEncryptTest#encrypt`
#[test]
fn poi_poi_encrypt_test_encrypt() {
    helpers::assert_encrypt();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiEncryptTest#encryptExcel`
#[test]
fn poi_poi_encrypt_test_encrypt_excel() {
    helpers::assert_encrypt();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiFormatTest#lastRowNum`
/// (was `D:\test\原文件.xlsx`)
#[test]
fn poi_poi_format_test_last_row_num() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiFormatTest#lastRowNumXSSF`
/// (was `/Users/.../测试格式.xlsx` — DataFormatter → dataformat fixture)
#[test]
fn poi_poi_format_test_last_row_num_xssf() {
    helpers::assert_dataformat_xlsx();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum`
/// (was `/Users/.../test3.xlsx`)
#[test]
fn poi_poi_test_last_row_num() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNumXSSF`
/// (XSSF style clone probe → fixture read smoke)
#[test]
fn poi_poi_test_last_row_num_xssf() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNumXSSFv22`
/// (HSSF style clone on local `.xls` → xls fixture)
#[test]
fn poi_poi_test_last_row_num_xss_fv_22() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum233`
/// (`TestFileUtil` + `fill/simple.xlsx`)
#[test]
fn poi_poi_test_last_row_num_233() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum255`
/// (was `D:\test\complex.xlsx` + shiftRows → `fill/complex.xlsx` read)
#[test]
fn poi_poi_test_last_row_num_255() {
    helpers::assert_fill_complex_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#cp`
/// (was `d://test/tt.xlsx`)
#[test]
fn poi_poi_test_cp() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum233443`
/// (was `d://test/em0.xlsx`)
#[test]
fn poi_poi_test_last_row_num_233443() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum2333`
#[test]
fn poi_poi_test_last_row_num_2333() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#testread`
#[test]
fn poi_poi_test_testread() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#testreadRead`
#[test]
fn poi_poi_test_testread_read() {
    helpers::assert_fill_simple_bytes();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum2332222`
#[test]
fn poi_poi_test_last_row_num_2332222() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum23443`
#[test]
fn poi_poi_test_last_row_num_23443() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNum2`
#[test]
fn poi_poi_test_last_row_num_2() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiTest#lastRowNumXSSF2`
#[test]
fn poi_poi_test_last_row_num_xssf_2() {
    helpers::assert_fill_simple_read();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write0`
#[test]
fn poi_poi_write_test_write_0() {
    helpers::assert_poi_long_write();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write01`
#[test]
fn poi_poi_write_test_write_01() {
    helpers::assert_float_decimal_smoke();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write`
#[test]
fn poi_poi_write_test_write() {
    helpers::assert_poi_long_write();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#write1`
#[test]
fn poi_poi_write_test_write_1() {
    helpers::assert_long2bytes();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#part`
#[test]
fn poi_poi_write_test_part() {
    helpers::assert_fill_placeholder_pattern();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#part2`
#[test]
fn poi_poi_write_test_part_2() {
    helpers::assert_fill_placeholder_pattern();
}

/// Java `com.alibaba.easyexcel.test.temp.poi.PoiWriteTest#part4`
/// (network image URL → `converter/img.jpg` fixture)
#[test]
fn poi_poi_write_test_part_4() {
    helpers::assert_image_write();
}
