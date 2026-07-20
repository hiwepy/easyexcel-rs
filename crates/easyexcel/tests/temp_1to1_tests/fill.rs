//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.fill.*`

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest#simpleFill`
#[test]
fn fill_fill_temp_test_simple_fill() {
    helpers::assert_fill_simple();
}

/// Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest#listFill`
#[test]
fn fill_fill_temp_test_list_fill() {
    helpers::assert_fill_list();
}

/// Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest#complexFill`
#[test]
fn fill_fill_temp_test_complex_fill() {
    helpers::assert_fill_complex();
}

/// Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest#complexFillWithTable`
#[test]
fn fill_fill_temp_test_complex_fill_with_table() {
    helpers::assert_fill_table();
}

/// Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest#horizontalFill`
#[test]
fn fill_fill_temp_test_horizontal_fill() {
    helpers::assert_fill_horizontal();
}

/// Java `com.alibaba.easyexcel.test.temp.fill.FillTempTest#compositeFill`
#[test]
fn fill_fill_temp_test_composite_fill() {
    helpers::assert_fill_composite();
}
