//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.large.*`

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#read`
#[test]
fn large_temp_large_data_test_read() {
    helpers::assert_large_fixture();
}

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#noModelRead`
#[test]
fn large_temp_large_data_test_no_model_read() {
    helpers::assert_no_model_read();
}

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#noModelRead2`
#[test]
fn large_temp_large_data_test_no_model_read_2() {
    helpers::assert_no_model_read();
}

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04Write`
#[test]
fn large_temp_large_data_test_t_04_write() {
    helpers::assert_large_write();
}

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04WriteExcel`
#[test]
fn large_temp_large_data_test_t_04_write_excel() {
    helpers::assert_large_write();
}

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04WriteExcelNo`
#[test]
fn large_temp_large_data_test_t_04_write_excel_no() {
    helpers::assert_large_write();
}

/// Java `com.alibaba.easyexcel.test.temp.large.TempLargeDataTest#t04WriteExcelPoi`
/// (SXSSFWorkbook large write → EasyExcel batched/large write)
#[test]
fn large_temp_large_data_test_t_04_write_excel_poi() {
    helpers::assert_large_write();
}
