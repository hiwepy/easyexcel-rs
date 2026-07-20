//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.write.*`

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.write.TempWriteTest#write`
#[test]
fn write_temp_write_test_write() {
    helpers::assert_write_newline();
}

/// Java `com.alibaba.easyexcel.test.temp.write.TempWriteTest#cglib`
#[test]
fn write_temp_write_test_cglib() {
    helpers::assert_write_newline();
}

/// Java `com.alibaba.easyexcel.test.temp.write.TempWriteTest#imageWrite`
#[test]
fn write_temp_write_test_image_write() {
    helpers::assert_image_write();
}

/// Java `com.alibaba.easyexcel.test.temp.write.TempWriteTest#imageWritePoi`
/// (SXSSF image write → EasyExcel image fixture + write smoke)
#[test]
fn write_temp_write_test_image_write_poi() {
    helpers::assert_image_write();
}

/// Java `com.alibaba.easyexcel.test.temp.write.TempWriteTest#tep`
/// (SXSSF image/tep → EasyExcel image fixture + write smoke)
#[test]
fn write_temp_write_test_tep() {
    helpers::assert_image_write();
}

/// Java `com.alibaba.easyexcel.test.temp.write.TempWriteTest#large`
#[test]
fn write_temp_write_test_large() {
    helpers::assert_large_write();
}
