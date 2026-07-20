//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.*`

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.FillTempTest#complexFill`
#[test]
fn fill_temp_test_complex_fill() {
    helpers::assert_fill_complex();
}

/// Java `com.alibaba.easyexcel.test.temp.FillTempTest#complexFillWithTable`
#[test]
fn fill_temp_test_complex_fill_with_table() {
    helpers::assert_fill_table();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#test`
#[test]
fn lock_2_test_test() {
    helpers::assert_converter_read();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#test33`
#[test]
fn lock_2_test_test_33() {
    helpers::assert_write_simple();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#write`
#[test]
fn lock_2_test_write() {
    helpers::assert_style_handler();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#simpleWrite`
#[test]
fn lock_2_test_simple_write() {
    helpers::assert_style_handler();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#testc`
#[test]
fn lock_2_test_testc() {
    helpers::assert_cell_reference_b3();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#simpleRead`
#[test]
fn lock_2_test_simple_read() {
    helpers::assert_head_read();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#test2`
#[test]
fn lock_2_test_test_2() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#test335`
#[test]
fn lock_2_test_test_335() {
    helpers::assert_cell_ref_parse();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt`
#[test]
fn lock_2_test_numberforamt() {
    helpers::assert_excel_serial_date_probes();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#testDate`
#[test]
fn lock_2_test_test_date() {
    helpers::assert_date_smoke();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#testDateAll`
#[test]
fn lock_2_test_test_date_all() {
    helpers::assert_excel_date_roundtrip_sampled();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt3`
#[test]
fn lock_2_test_numberforamt_3() {
    helpers::assert_dataformat_dates();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt4`
#[test]
fn lock_2_test_numberforamt_4() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt77`
#[test]
fn lock_2_test_numberforamt_77() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt99`
#[test]
fn lock_2_test_numberforamt_99() {
    helpers::assert_datetime_nanos_format();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt5`
#[test]
fn lock_2_test_numberforamt_5() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt6`
#[test]
fn lock_2_test_numberforamt_6() {
    helpers::assert_decimal_scale_smoke();
}

/// Java `com.alibaba.easyexcel.test.temp.Lock2Test#numberforamt7`
#[test]
fn lock_2_test_numberforamt_7() {
    helpers::assert_decimal_scale_smoke();
}

/// Java `com.alibaba.easyexcel.test.temp.LockTest#test`
#[test]
fn lock_test_test() {
    helpers::assert_head_read();
}

/// Java `com.alibaba.easyexcel.test.temp.LockTest#test2`
#[test]
fn lock_test_test_2() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#test`
#[test]
fn style_test_test() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#poi`
#[test]
fn style_test_poi() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#poi07`
#[test]
fn style_test_poi_07() {
    helpers::assert_dataformat_xlsx();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#poi0701`
#[test]
fn style_test_poi_0701() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#poi0702`
#[test]
fn style_test_poi_0702() {
    helpers::assert_dataformat_xlsx();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#poi0703`
#[test]
fn style_test_poi_0703() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#testFormatter`
#[test]
fn style_test_test_formatter() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#testFormatter2`
#[test]
fn style_test_test_formatter_2() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#testFormatter3`
#[test]
fn style_test_test_formatter_3() {
    helpers::assert_style_format();
}

/// Java `com.alibaba.easyexcel.test.temp.StyleTest#testBuiltinFormats`
#[test]
fn style_test_test_builtin_formats() {
    helpers::assert_style_write();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteLargeTest#test`
#[test]
fn write_large_test_test() {
    helpers::assert_large_write();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteLargeTest#read`
#[test]
fn write_large_test_read() {
    helpers::assert_large_fixture();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteLargeTest#read2`
#[test]
fn write_large_test_read_2() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteLargeTest#read3`
#[test]
fn write_large_test_read_3() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteLargeTest#test2`
#[test]
fn write_large_test_test_2() {
    helpers::assert_large_batched();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteV33Test#handlerStyleWrite`
#[test]
fn write_v_33_test_handler_style_write() {
    helpers::assert_style_handler();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteV33Test#test4`
#[test]
fn write_v_33_test_test_4() {
    helpers::assert_style_handler();
}

/// Java `com.alibaba.easyexcel.test.temp.WriteV34Test#test`
#[test]
fn write_v_34_test_test() {
    helpers::assert_style_handler();
}

/// Java `com.alibaba.easyexcel.test.temp.Xls03Test#test`
#[test]
fn xls_03_test_test() {
    helpers::assert_xls_read();
}

/// Java `com.alibaba.easyexcel.test.temp.Xls03Test#test2`
#[test]
fn xls_03_test_test_2() {
    helpers::assert_xls_read();
}
